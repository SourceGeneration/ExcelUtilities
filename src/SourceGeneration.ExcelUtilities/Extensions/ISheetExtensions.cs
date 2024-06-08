using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace SourceGeneration.ExcelUtilities;

internal static class ISheetExtensions
{
    public static List<T> ReadRows<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(this ISheet sheet, ExcelLoadOptions? options = null) where T : class, new()
    {
        options ??= ExcelLoadOptions.Default;

        var columns = options.GetColumns(typeof(T));

        var records = new List<T>();
        for (int r = options.StartRow; r <= sheet.LastRowNum; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null)
                continue;

            var obj = new T();

            int i = 0;
            foreach (var column in columns)
            {
                int cellNumber = 0;

                if (column.Order == -1)
                    cellNumber = i + options.StartColumn;
                else
                    cellNumber = column.Order + options.StartColumn;

                i++;

                var cell = row.GetCell(cellNumber);
                if (cell == null)
                    continue;

                column.ReadCell(cell, obj);
            }

            records.Add(obj);
        }
        return records;
    }

    public static void WriteRows<T>(this ISheet sheet,
        int row,
        IEnumerable<T> items,
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    Type itemType,
        ExcelTableStyles styles,
        ExcelSaveOptions options)
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));

        var columns = options.GetColumns(itemType);

        if (options.AutoCreateHeader)
            WriteHeaders(sheet, ref row, columns, styles);

        WriteRowsCore(sheet, row, items, columns, styles, options);
    }

    private static void WriteRowsCore<T>(ISheet sheet, int rowNumber, IEnumerable<T> items, IReadOnlyList<ExcelColumnBase> columns, ExcelTableStyles styles, ExcelSaveOptions options)
    {
        var start = rowNumber - 1;

        foreach (var item in items)
        {
            if (item == null)
                continue;

            var row = sheet.CreateRow(rowNumber);

            if (options.AutoSizeRow)
                row.HeightInPoints = -1;
            else
                row.HeightInPoints = options.RowHeight;

            int col = 0;

            foreach (var column in columns)
            {
                var cell = row.CreateCell(col);
                column.WriteCell(cell, item, rowNumber - start, styles.DataStyleTable, options);
                col++;
            }

            rowNumber++;
        }

        if (options.AutoSizeColumn)
        {
            int col = 0;
            foreach (var column in columns)
            {
                var size = column.ComputeColumnSize();
                if (size == 0)
                {
                    try
                    {
                        sheet.AutoSizeColumn(col);
                    }
                    catch
                    {
                        //当字符串中有特殊字符时，会产生错误
                        sheet.SetColumnWidth(col, 256 * 40);
                    }
                }
                else
                    sheet.SetColumnWidth(col, size);
                col++;
            }
        }
    }

    private static void WriteHeaders(ISheet sheet, ref int row, IReadOnlyList<ExcelColumnBase> columns, ExcelTableStyles styles)
    {
        var header_row = sheet.CreateRow(row);
        header_row.HeightInPoints = 20;
        row++;
        int i = 0;

        for (int colIndex = 0; colIndex < columns.Count; colIndex++)
        {
            ExcelColumnBase? column = columns[colIndex];
            if (column.DataType == ExcelColumnDataType.String)
                sheet.SetDefaultColumnStyle(colIndex, styles.DataStyleTable.DefaultTextCellStyle);
        }

        foreach (var column in columns)
        {
            var cell = header_row.CreateCell(i);
            cell.SetCellType(CellType.String);

            if (column is ExcelRowNumberColumn)
            {
                cell.CellStyle = styles.HeaderStyleTable.AlignCenterStyle;
            }
            else if (column.DataType == ExcelColumnDataType.Boolean ||
                column.DataType == ExcelColumnDataType.Date ||
                column.DataType == ExcelColumnDataType.Time ||
                column.DataType == ExcelColumnDataType.DateTime ||
                column.DataType == ExcelColumnDataType.Enum)
            {
                cell.CellStyle = styles.HeaderStyleTable.AlignCenterStyle;
            }
            else if (column.DataType == ExcelColumnDataType.Integer || column.DataType == ExcelColumnDataType.Number)
            {
                cell.CellStyle = styles.HeaderStyleTable.AlignRightStyle;
            }
            else
            {
                cell.CellStyle = styles.HeaderStyleTable.AlignLeftStyle;
            }

            cell.SetCellValue(column.Title);
            i++;
        }
    }

}