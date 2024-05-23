namespace SourceGeneration.ExcelUtilities;

/// <inheritdoc/>
public class ExcelLoadOptions : ExcelOptions
{
    public static readonly ExcelLoadOptions Default = new();

    public int StartRow { get; set; } = 0;

    public int StartColumn { get; set; } = 0;
}
