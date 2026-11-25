using System.Buffers.Binary;

namespace Avig.OpenXml;

/// <summary>
/// Reads a zip file's directory on its own and checks that the file says one thing only. A zip
/// file names its entries twice, in a local header before each entry's data and in the central
/// directory at the end, and readers differ in which one they believe: .NET and Office read the
/// directory, many scanners and unpackers walk the local headers. A file where the two disagree,
/// where entries overlap, or with bytes before, between or after the entries can show one
/// program one document and another program another, which is how files are made to look
/// harmless to a scanner. Office, LibreOffice and the .NET writers leave no such room.
/// </summary>
internal static class ZipStructure
{
    private const uint LocalHeaderSignature = 0x04034b50;
    private const uint CentralHeaderSignature = 0x02014b50;
    private const uint EndSignature = 0x06054b50;
    private const uint Zip64EndSignature = 0x06064b50;
    private const uint Zip64LocatorSignature = 0x07064b50;
    private const uint DataDescriptorSignature = 0x08074b50;
    private const int EndLength = 22;
    private const int Zip64LocatorLength = 20;
    private const int CentralHeaderLength = 46;
    private const int LocalHeaderLength = 30;

    public static async Task<Result> CheckAsync(Stream zip, int maxEntries, CancellationToken cancellationToken)
    {
        long start = zip.Position;

        try
        {
            return await CheckFromAsync(zip, start, zip.Length - start, maxEntries, cancellationToken);
        }
        catch (EndOfStreamException)
        {
            return Damaged("an entry or record runs past the end of the file");
        }
        finally
        {
            zip.Position = start;
        }
    }

    private static async Task<Result> CheckFromAsync(Stream zip, long start, long length, int maxEntries, CancellationToken cancellationToken)
    {
        if (length < EndLength)
        {
            return Damaged("it is too short to be a zip file");
        }

        // The end record is the last thing in the file, followed only by its comment: the comment
        // length has to reach the end exactly, so that nothing is hidden after it.
        int tailLength = (int)Math.Min(length, EndLength + ushort.MaxValue);
        byte[] tail = await ReadAsync(zip, start + length - tailLength, tailLength, cancellationToken);
        int end = -1;

        for (int position = tailLength - EndLength; position >= 0; position--)
        {
            if (BinaryPrimitives.ReadUInt32LittleEndian(tail.AsSpan(position)) == EndSignature
                && BinaryPrimitives.ReadUInt16LittleEndian(tail.AsSpan(position + 20)) == tailLength - position - EndLength)
            {
                end = position;

                break;
            }
        }

        if (end < 0)
        {
            return Damaged("its end record is missing, or there are bytes after it");
        }

        long endOffset = length - tailLength + end;
        ReadOnlySpan<byte> record = tail.AsSpan(end);

        if (BinaryPrimitives.ReadUInt16LittleEndian(record[4..]) != 0 || BinaryPrimitives.ReadUInt16LittleEndian(record[6..]) != 0)
        {
            return Damaged("it is split over several disks");
        }

        long entries = BinaryPrimitives.ReadUInt16LittleEndian(record[10..]);
        long directorySize = BinaryPrimitives.ReadUInt32LittleEndian(record[12..]);
        long directoryOffset = BinaryPrimitives.ReadUInt32LittleEndian(record[16..]);
        long directoryEnd = endOffset;

        if (entries != BinaryPrimitives.ReadUInt16LittleEndian(record[8..]))
        {
            return Damaged("its end record gives two different numbers of entries");
        }

        bool zip64 = entries == ushort.MaxValue || directorySize == uint.MaxValue || directoryOffset == uint.MaxValue;

        if (zip64 == true || (endOffset >= Zip64LocatorLength && await SignatureAtAsync(zip, start + endOffset - Zip64LocatorLength, Zip64LocatorSignature, cancellationToken) == true))
        {
            Result<(long Entries, long Size, long Offset, long RecordOffset)> large = await ReadZip64EndAsync(zip, start, endOffset, cancellationToken);

            if (large.IsFailure(out string? error, out _, out (long Entries, long Size, long Offset, long RecordOffset) values) == true)
            {
                return Result.Invalid(error);
            }

            (entries, directorySize, directoryOffset, directoryEnd) = values;
        }

        if (entries > maxEntries)
        {
            // The limit on parts is checked, with its own message, once the archive is open.
            return Result.Success();
        }

        if (directoryOffset < 0 || directorySize < 0 || directoryOffset + directorySize != directoryEnd)
        {
            return Damaged("its central directory does not end where the end record begins");
        }

        if (directorySize > (long)entries * (CentralHeaderLength + (3 * ushort.MaxValue)))
        {
            return Damaged("its central directory is larger than its entries can be");
        }

        byte[] directory = await ReadAsync(zip, start + directoryOffset, (int)directorySize, cancellationToken);
        List<(long Start, long End, string Name)> spans = [];
        int at = 0;

        for (long index = 0; index < entries; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (at + CentralHeaderLength > directory.Length || BinaryPrimitives.ReadUInt32LittleEndian(directory.AsSpan(at)) != CentralHeaderSignature)
            {
                return Damaged("its central directory has fewer entries than its end record says");
            }

            ReadOnlySpan<byte> header = directory.AsSpan(at);
            ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(header[8..]);
            ushort method = BinaryPrimitives.ReadUInt16LittleEndian(header[10..]);
            long compressed = BinaryPrimitives.ReadUInt32LittleEndian(header[20..]);
            long uncompressed = BinaryPrimitives.ReadUInt32LittleEndian(header[24..]);
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
            int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[30..]);
            int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
            long localOffset = BinaryPrimitives.ReadUInt32LittleEndian(header[42..]);
            int next = at + CentralHeaderLength + nameLength + extraLength + commentLength;

            if (next > directory.Length)
            {
                return Damaged("an entry in its central directory runs past the directory's end");
            }

            byte[] name = directory.AsSpan(at + CentralHeaderLength, nameLength).ToArray();
            string shown = System.Text.Encoding.UTF8.GetString(name);
            ReadOnlySpan<byte> extra = directory.AsSpan(at + CentralHeaderLength + nameLength, extraLength);

            if (TryApplyZip64(extra, ref uncompressed, ref compressed, ref localOffset) == false)
            {
                return Damaged($"the entry '{shown}' has a broken zip64 field");
            }

            Result<long> entryEnd = await CheckLocalAsync(zip, start, directoryOffset, localOffset, name, shown, flags, method, compressed, uncompressed, cancellationToken);

            if (entryEnd.IsFailure(out string? error, out _, out long stop) == true)
            {
                return Result.Invalid(error);
            }

            spans.Add((localOffset, stop, shown));
            at = next;
        }

        if (at != directory.Length)
        {
            return Damaged("its central directory has more in it than its end record says");
        }

        // The entries have to fill the file from its first byte to the central directory, one
        // after another: no bytes before them, between them or after them, and no two sharing bytes.
        spans.Sort((left, right) => left.Start.CompareTo(right.Start));
        long expected = 0;

        foreach ((long spanStart, long spanEnd, string name) in spans)
        {
            if (spanStart < expected)
            {
                return Damaged($"the entry '{name}' overlaps the one before it");
            }

            if (spanStart > expected)
            {
                return Damaged($"there are {spanStart - expected:N0} bytes before the entry '{name}' that belong to no entry");
            }

            expected = spanEnd;
        }

        if (expected != directoryOffset)
        {
            return Damaged($"there are {directoryOffset - expected:N0} bytes between the last entry and the central directory");
        }

        return Result.Success();
    }

    /// <summary>The local header has to repeat the directory's name, method and flags; the result is where the entry, with its data descriptor, ends.</summary>
    private static async Task<Result<long>> CheckLocalAsync(Stream zip, long start, long directoryOffset, long localOffset, byte[] name, string shown, ushort flags, ushort method, long compressed, long uncompressed, CancellationToken cancellationToken)
    {
        if (localOffset < 0 || localOffset + LocalHeaderLength > directoryOffset)
        {
            return Result<long>.Invalid(DamagedMessage($"the entry '{shown}' starts outside the file's data"));
        }

        byte[] local = await ReadAsync(zip, start + localOffset, LocalHeaderLength, cancellationToken);
        ReadOnlySpan<byte> header = local;

        if (BinaryPrimitives.ReadUInt32LittleEndian(header) != LocalHeaderSignature)
        {
            return Result<long>.Invalid(DamagedMessage($"the entry '{shown}' has no local header where the directory says"));
        }

        ushort localFlags = BinaryPrimitives.ReadUInt16LittleEndian(header[6..]);
        ushort localMethod = BinaryPrimitives.ReadUInt16LittleEndian(header[8..]);
        int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
        int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);

        // Bit 0 is encryption and bit 3 the data descriptor; the rest (UTF-8 names, compression
        // options) may differ between the two headers in files that open everywhere.
        const ushort Significant = 0x0009;

        if (localMethod != method || (localFlags & Significant) != (flags & Significant))
        {
            return Result<long>.Invalid(DamagedMessage($"the entry '{shown}' has another compression or encryption in its local header than in the directory"));
        }

        if (localOffset + LocalHeaderLength + nameLength + extraLength > directoryOffset)
        {
            return Result<long>.Invalid(DamagedMessage($"the local header of '{shown}' runs into the central directory"));
        }

        byte[] localName = await ReadAsync(zip, start + localOffset + LocalHeaderLength, nameLength, cancellationToken);

        if (localName.AsSpan().SequenceEqual(name) == false)
        {
            return Result<long>.Invalid(DamagedMessage($"the entry '{shown}' has another name in its local header than in the directory"));
        }

        long dataEnd = localOffset + LocalHeaderLength + nameLength + extraLength + compressed;

        if (dataEnd > directoryOffset)
        {
            return Result<long>.Invalid(DamagedMessage($"the data of '{shown}' runs into the central directory"));
        }

        if ((flags & 0x0008) == 0)
        {
            return dataEnd;
        }

        // A data descriptor follows the data: an optional signature, the CRC and the two sizes,
        // which are 8 bytes each in zip64. Of the lengths it can have, one must fit.
        bool large = compressed >= uint.MaxValue || uncompressed >= uint.MaxValue;
        int sizes = 8;

        if (large == true)
        {
            sizes = 16;
        }

        bool signed = dataEnd + 4 <= directoryOffset && await SignatureAtAsync(zip, start + dataEnd, DataDescriptorSignature, cancellationToken) == true;
        long descriptor = 4 + sizes;

        if (signed == true)
        {
            descriptor += 4;
        }

        if (dataEnd + descriptor > directoryOffset)
        {
            return Result<long>.Invalid(DamagedMessage($"the data descriptor of '{shown}' runs into the central directory"));
        }

        return dataEnd + descriptor;
    }

    /// <summary>The zip64 end record, through its locator just before the end record.</summary>
    private static async Task<Result<(long Entries, long Size, long Offset, long RecordOffset)>> ReadZip64EndAsync(Stream zip, long start, long endOffset, CancellationToken cancellationToken)
    {
        if (endOffset < Zip64LocatorLength)
        {
            return Result<(long, long, long, long)>.Invalid(DamagedMessage("its zip64 locator is missing"));
        }

        byte[] locator = await ReadAsync(zip, start + endOffset - Zip64LocatorLength, Zip64LocatorLength, cancellationToken);

        if (BinaryPrimitives.ReadUInt32LittleEndian(locator) != Zip64LocatorSignature)
        {
            return Result<(long, long, long, long)>.Invalid(DamagedMessage("its zip64 locator is missing"));
        }

        long recordOffset = (long)BinaryPrimitives.ReadUInt64LittleEndian(locator.AsSpan(8));

        // The zip64 end record (56 bytes, with no extensible data in files anyone writes) has to
        // end where the locator begins.
        if (recordOffset < 0 || recordOffset + 56 != endOffset - Zip64LocatorLength)
        {
            return Result<(long, long, long, long)>.Invalid(DamagedMessage("its zip64 end record is not where its locator says"));
        }

        byte[] record = await ReadAsync(zip, start + recordOffset, 56, cancellationToken);

        if (BinaryPrimitives.ReadUInt32LittleEndian(record) != Zip64EndSignature)
        {
            return Result<(long, long, long, long)>.Invalid(DamagedMessage("its zip64 end record is missing"));
        }

        long entries = (long)BinaryPrimitives.ReadUInt64LittleEndian(record.AsSpan(32));
        long size = (long)BinaryPrimitives.ReadUInt64LittleEndian(record.AsSpan(40));
        long offset = (long)BinaryPrimitives.ReadUInt64LittleEndian(record.AsSpan(48));

        if (entries < 0 || entries != (long)BinaryPrimitives.ReadUInt64LittleEndian(record.AsSpan(24)))
        {
            return Result<(long, long, long, long)>.Invalid(DamagedMessage("its zip64 end record gives two different numbers of entries"));
        }

        return (entries, size, offset, recordOffset);
    }

    /// <summary>Replaces the fields set to 0xFFFFFFFF with the zip64 extra field's, in its order: uncompressed, compressed, offset.</summary>
    private static bool TryApplyZip64(ReadOnlySpan<byte> extra, ref long uncompressed, ref long compressed, ref long localOffset)
    {
        int at = 0;

        while (at + 4 <= extra.Length)
        {
            ushort id = BinaryPrimitives.ReadUInt16LittleEndian(extra[at..]);
            int size = BinaryPrimitives.ReadUInt16LittleEndian(extra[(at + 2)..]);

            if (at + 4 + size > extra.Length)
            {
                return false;
            }

            if (id == 0x0001)
            {
                ReadOnlySpan<byte> field = extra.Slice(at + 4, size);
                int read = 0;

                if (TryTake(field, ref read, uncompressed, out uncompressed) == false || TryTake(field, ref read, compressed, out compressed) == false || TryTake(field, ref read, localOffset, out localOffset) == false)
                {
                    return false;
                }
            }

            at += 4 + size;
        }

        return true;
    }

    private static bool TryTake(ReadOnlySpan<byte> field, ref int read, long value, out long result)
    {
        result = value;

        if (value != uint.MaxValue)
        {
            return true;
        }

        if (read + 8 > field.Length)
        {
            return false;
        }

        result = (long)BinaryPrimitives.ReadUInt64LittleEndian(field[read..]);
        read += 8;

        return result >= 0;
    }

    private static async Task<bool> SignatureAtAsync(Stream zip, long position, uint signature, CancellationToken cancellationToken)
    {
        if (position < 0 || position + 4 > zip.Length)
        {
            return false;
        }

        byte[] bytes = await ReadAsync(zip, position, 4, cancellationToken);

        return BinaryPrimitives.ReadUInt32LittleEndian(bytes) == signature;
    }

    private static async Task<byte[]> ReadAsync(Stream zip, long position, int count, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[count];
        zip.Position = position;
        await zip.ReadExactlyAsync(buffer, cancellationToken);

        return buffer;
    }

    private static Result Damaged(string reason)
    {
        return Result.Invalid(DamagedMessage(reason));
    }

    private static string DamagedMessage(string reason)
    {
        return $"{OpcErrors.DamagedPrefix}: the zip file does not say one thing only, {reason}.";
    }
}
