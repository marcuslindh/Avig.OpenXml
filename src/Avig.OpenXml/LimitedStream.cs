namespace Avig.OpenXml;

/// <summary>
/// Reads a zip entry and refuses to give more bytes than its header said it holds. The total
/// size check when the file is opened adds up the headers, and without this a part could claim
/// a kilobyte and unpack to gigabytes.
/// </summary>
internal sealed class LimitedStream : Stream
{
    private readonly Stream inner;
    private readonly long limit;
    private readonly string name;
    private long read;

    public LimitedStream(Stream inner, long limit, string name)
    {
        this.inner = inner;
        this.limit = limit;
        this.name = name;
    }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => read;
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        return Count(inner.Read(buffer, offset, count));
    }

    public override int Read(Span<byte> buffer)
    {
        return Count(inner.Read(buffer));
    }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        return Count(await inner.ReadAsync(buffer.AsMemory(offset, count), cancellationToken));
    }

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        return Count(await inner.ReadAsync(buffer, cancellationToken));
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        throw new NotSupportedException();
    }

    public override void SetLength(long value)
    {
        throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotSupportedException();
    }

    public override async ValueTask DisposeAsync()
    {
        await inner.DisposeAsync();
        await base.DisposeAsync();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing == true)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
    }

    private int Count(int bytes)
    {
        read += bytes;

        if (read > limit)
        {
            throw new InvalidDataException($"The file is damaged: the part '{name}' unpacks to more than the {limit} bytes its zip header gives.");
        }

        return bytes;
    }
}
