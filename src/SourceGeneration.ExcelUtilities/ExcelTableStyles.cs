using NPOI.SS.UserModel;

namespace SourceGeneration.ExcelUtilities;

public class ExcelTableStyles
{
    public ExcelRowStyles DataStyleTable { get; init; } = default!;
    public ExcelHeaderStyles HeaderStyleTable { get; init; } = default!;
}

public class ExcelRowStyles
{
    public IFont Font { get; init; } = default!;
    public ICellStyle DateTimeCellStyle { get; init; } = default!;
    public ICellStyle DateCellStyle { get; init; } = default!;
    public ICellStyle TimeCellStyle { get; init; } = default!;
    public ICellStyle NumberCellStyle { get; init; } = default!;
    public ICellStyle IntegerCellStyle { get; init; } = default!;
    public ICellStyle TextCellStyle { get; init; } = default!;
    public ICellStyle CenterCellStyle { get; init; } = default!;
    public ICellStyle LeftCellStyle { get; init; } = default!;
    public ICellStyle DefaultTextCellStyle { get; init; } = default!;
}

public class ExcelHeaderStyles
{
    public IFont Font { get; init; } = default!;
    public ICellStyle AlignLeftStyle { get; init; } = default!;
    public ICellStyle AlignCenterStyle { get; init; } = default!;
    public ICellStyle AlignRightStyle { get; init; } = default!;
}

