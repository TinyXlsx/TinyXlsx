namespace TinyXlsx;

/// <summary>
/// Represents a stylesheet for an XLSX document.
/// </summary>
public class Stylesheet
{
    /// <summary>
    /// The formats which are part of the stylesheet.
    /// </summary>
    public Dictionary<string, (int ZeroBasedIndex, int CustomFormatIndex)> Formats { get; init; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Stylesheet"/> class.
    /// </summary>
    public Stylesheet()
    {
        Formats = [];
    }

    /// <summary>
    /// Gets or creates a unique number format style for the specified format string.
    /// </summary>
    /// <param name="format">
    /// The format string to get or create.
    /// </param>
    /// <returns>
    /// A tuple containing the zero-based index and custom format index for the style.
    /// </returns>
    /// <exception cref="NotSupportedException">
    /// Thrown if the number of styles exceeds the maximum supported by the XLSX format.
    /// </exception>
    public (int ZeroBasedIndex, int CustomFormatIndex) GetOrCreateNumberFormat(string format)
    {
        var count = Formats.Count;

        if (count >= Constants.MaximumStyles)
        {
            throw new NotSupportedException("The XLSX format does not support more than 65,490 styles.");
        }

        if (Formats.TryGetValue(format, out var indexes))
        {
            return indexes;
        }

        indexes = (count + 1, count + 164);
        Formats.Add(format, indexes);
        return indexes;
    }
}
