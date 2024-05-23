using NPOI.SS.UserModel;

namespace SourceGeneration.ExcelUtilities;

public abstract class ExcelColumnBase
{
    public string Name { get; protected set; } = null!;
    public string? DisplayName { get; protected set; }
    public int Order { get; protected set; }
    public ExcelColumnDataType ValueType { get; protected set; }

    public abstract void ReadCell(ICell cell, object obj);
    public abstract void WriteCell(ICell cell, object obj, int rowNumber, ExcelRowStyles styles, ExcelSaveOptions options);

    public virtual int ComputeColumnSize() => 0;
}
