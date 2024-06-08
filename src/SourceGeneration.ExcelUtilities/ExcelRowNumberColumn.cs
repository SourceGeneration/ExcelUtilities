using NPOI.SS.UserModel;

namespace SourceGeneration.ExcelUtilities;

/// <inheritdoc/>
public sealed class ExcelRowNumberColumn : ExcelColumnBase
{
    /// <inheritdoc/>
    public ExcelRowNumberColumn(string title)
    {
        Title = title;
        Order = 0;
        DataType = ExcelColumnDataType.Integer;
    }

    /// <inheritdoc/>
    public override void ReadCell(ICell cell, object obj) { }

    /// <inheritdoc/>
    public override void WriteCell(ICell cell, object obj, int rowNumber, ExcelRowStyles styles, ExcelSaveOptions options)
    {
        cell.SetCellType(CellType.Numeric);
        cell.CellStyle = styles.CenterCellStyle;
        cell.SetCellValue(rowNumber);
    }
}