using NPOI.SS.UserModel;

namespace SourceGeneration.ExcelUtilities;

public abstract class ExcelColumnBase
{
    public string? Title { get; protected set; }
    public int Order { get; protected set; }
    public ExcelColumnDataType DataType { get; protected set; }

    public abstract void ReadCell(ICell cell, object obj);
    public abstract void WriteCell(ICell cell, object obj, int rowNumber, ExcelRowStyles styles, ExcelSaveOptions options);

    public virtual int ComputeColumnSize() => 0;
}
