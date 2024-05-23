using NPOI.SS.UserModel;

namespace SourceGeneration.ExcelUtilities;

internal static class IWorkbookStyleExtensions
{
    public static ExcelTableStyles CreateDefaultStyleTable(this IWorkbook workbook)
    {
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        if (System.Globalization.CultureInfo.CurrentUICulture.Name == "zh-CN")
            headerFont.FontName = "宋体";
        headerFont.FontHeightInPoints = 12;

        var cellFont = workbook.CreateFont();
        cellFont.IsBold = false;
        if (System.Globalization.CultureInfo.CurrentUICulture.Name == "zh-CN")
            cellFont.FontName = "宋体";
        cellFont.FontHeightInPoints = 11;

        return new ExcelTableStyles
        {
            DataStyleTable = new ExcelRowStyles
            {
                Font = cellFont,
                DateCellStyle = CreateDateCellStyle(workbook, cellFont),
                DateTimeCellStyle = CreateDateTimeCellStyle(workbook, cellFont),
                TimeCellStyle = CreateTimeCellStyle(workbook, cellFont),
                IntegerCellStyle = CreateIntegerCellStyle(workbook, cellFont),
                NumberCellStyle = CreateNumberCellStyle(workbook, cellFont),
                TextCellStyle = CreateTextCellStyle(workbook, cellFont),
                CenterCellStyle = CreateCellStyle(workbook, cellFont, HorizontalAlignment.Center),
                LeftCellStyle = CreateCellStyle(workbook, cellFont, HorizontalAlignment.Left),
                DefaultTextCellStyle  = CreateDefaultTextCellStyle(workbook, cellFont),
            },
            HeaderStyleTable = new ExcelHeaderStyles
            {
                Font = headerFont,
                AlignLeftStyle = CreateHeaderCellStyle(workbook, headerFont, HorizontalAlignment.Left),
                AlignRightStyle = CreateHeaderCellStyle(workbook, headerFont, HorizontalAlignment.Right),
                AlignCenterStyle = CreateHeaderCellStyle(workbook, headerFont, HorizontalAlignment.Center),
            }
        };
    }

    private static ICellStyle CreateDateCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("m/d/yy");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Center, id);
    }

    private static ICellStyle CreateDateTimeCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("yyyy/m/d\\ h:mm;@");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Center, id);
    }

    private static ICellStyle CreateTimeCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("h:mm:ss;@");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Center, id);
    }

    private static ICellStyle CreateIntegerCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("0");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Right, id);
    }

    private static ICellStyle CreateNumberCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("0.00");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Right, id);
    }

    private static ICellStyle CreateTextCellStyle(IWorkbook workbook, IFont font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var id = format.GetFormat("@");
        return CreateCellStyle(workbook, font, HorizontalAlignment.Left, id);
    }


    private static ICellStyle CreateDefaultTextCellStyle(IWorkbook workbook, IFont? font)
    {
        IDataFormat format = workbook.CreateDataFormat();
        var cellStyle = workbook.CreateCellStyle();
        if (font != null)
        {
            cellStyle.SetFont(font);
        }
        cellStyle.DataFormat = format.GetFormat("@"); ;
        return cellStyle;
    }

    private static ICellStyle CreateCellStyle(IWorkbook workbook,
        IFont font,
        HorizontalAlignment alignment,
        short format = -1)
    {
        var cellStyle = workbook.CreateCellStyle();
        cellStyle.SetFont(font);
        cellStyle.Alignment = alignment;
        cellStyle.VerticalAlignment = VerticalAlignment.Center;
        cellStyle.BorderLeft = BorderStyle.Thin;
        cellStyle.BorderRight = BorderStyle.Thin;
        cellStyle.BorderTop = BorderStyle.Thin;
        cellStyle.BorderBottom = BorderStyle.Thin;
        cellStyle.DataFormat = format > 0 ? format : (short)0;
        
        return cellStyle;
    }

    private static ICellStyle CreateHeaderCellStyle(IWorkbook workbook,
        IFont font,
        HorizontalAlignment alignment)
    {
        var cellStyle = workbook.CreateCellStyle();
        cellStyle.SetFont(font);
        cellStyle.Alignment = alignment;
        cellStyle.VerticalAlignment = VerticalAlignment.Center;
        cellStyle.BorderLeft = BorderStyle.Thin;
        cellStyle.BorderRight = BorderStyle.Thin;
        cellStyle.BorderTop = BorderStyle.Thin;
        cellStyle.BorderBottom = BorderStyle.Thin;
        //cellStyle.DataFormat = format > 0 ? format : (short)0;

        //if (cellStyle is XSSFCellStyle xssfStyle)
        //    xssfStyle.SetFillForegroundColor(ForegroundColor);

        cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.PaleBlue.Index;
        cellStyle.FillPattern = FillPattern.SolidForeground;

        return cellStyle;
    }
}
