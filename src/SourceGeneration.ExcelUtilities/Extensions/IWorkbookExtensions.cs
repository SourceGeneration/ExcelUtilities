using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace SourceGeneration.ExcelUtilities;

internal static class IWorkbookExtensions
{
    public static void Save<T>(this IWorkbook workbook,
                               string filename,
                               string sheet,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                                Type itemType,
                               ExcelSaveOptions? options = null) => Save(workbook, filename, new Dictionary<string, IEnumerable<T>> { { sheet, rows } }, itemType, options);

    public static void Save<T>(this IWorkbook workbook,
                               string filename,
                               IDictionary<string, IEnumerable<T>> sheets,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        using var stream = CreateFile(filename);
        Save(workbook, stream, sheets, itemType, options);

        static FileStream CreateFile(string filename)
        {
            var dir = Path.GetDirectoryName(filename);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return File.Create(filename);
        }
    }

    public static void Save<T>(this IWorkbook workbook,
                               Stream stream,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null) => Save<T>(workbook, stream, new Dictionary<string, IEnumerable<T>> { { ExcelConstants.DefaultSheetName, rows } }, itemType, options);

    public static void Save<T>(this IWorkbook workbook,
                               Stream stream,
                               string sheet,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null) => Save<T>(workbook, stream, new Dictionary<string, IEnumerable<T>> { { sheet, rows } }, itemType, options);

    public static void Save<T>(this IWorkbook workbook,
                               Stream stream,
                               IDictionary<string, IEnumerable<T>> sheets,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        options ??= ExcelSaveOptions.Default;

        var styles = workbook.CreateDefaultStyleTable();

        foreach (var sheet in sheets)
        {
            workbook.CreateSheet(sheet.Key).WriteRows(0, sheet.Value, itemType, styles, options);
        }

        if (workbook is XSSFWorkbook xssf)
            xssf.Write(stream, true);
        else
            workbook.Write(stream);
    }
}
