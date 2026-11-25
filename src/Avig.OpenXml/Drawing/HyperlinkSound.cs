namespace Avig.OpenXml.Drawing;

/// <summary>A sound a click on a link plays (a:snd): a WAV file and the name shown for it.</summary>
public sealed record HyperlinkSound
{
    private HyperlinkSound(string name, ReadOnlyMemory<byte> data, bool check)
    {
        if (check == true)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            DrawingValues.XmlText(name, nameof(name));

            if (data.IsEmpty == true)
            {
                throw new ArgumentException("A sound has data.", nameof(data));
            }
        }

        Name = name;
        Data = data;
    }

    /// <summary>The name shown for the sound.</summary>
    public string Name { get; }

    /// <summary>The WAV file's bytes; empty for a sound read from a file whose data was not loaded.</summary>
    public ReadOnlyMemory<byte> Data { get; }

    /// <summary>A sound from a WAV file.</summary>
    /// <param name="name">The name shown for the sound, such as the file's name.</param>
    /// <param name="wav">The WAV file's bytes.</param>
    /// <returns>The sound.</returns>
    public static HyperlinkSound FromWav(string name, ReadOnlyMemory<byte> wav)
    {
        return new HyperlinkSound(name, wav, true);
    }

    /// <summary>A sound read from a file, which may come without its data. For readers.</summary>
    /// <param name="name">The name.</param>
    /// <param name="data">The bytes, or empty.</param>
    /// <returns>The sound.</returns>
    public static HyperlinkSound Read(string name, ReadOnlyMemory<byte> data)
    {
        ArgumentNullException.ThrowIfNull(name);

        return new HyperlinkSound(name, data, false);
    }
}
