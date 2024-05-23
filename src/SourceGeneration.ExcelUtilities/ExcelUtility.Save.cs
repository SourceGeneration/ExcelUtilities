using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;

namespace SourceGeneration.ExcelUtilities;

public static partial class ExcelUtility
{
    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, IEnumerable<T> rows, int perSheetRowCount, ExcelSaveOptions? options = null)
    {
        if (perSheetRowCount > 0)
        {
            int current = 0;
            int pageIndex = 1;
            Dictionary<string, IEnumerable<T>> pages = [];

            while (true)
            {
                var pageRows = rows.Skip(current).Take(perSheetRowCount).ToList();
                if (pageRows.Count == 0)
                    break;
                pages.Add(pageIndex.ToString(), pageRows);

                if (pageRows.Count < perSheetRowCount)
                    break;

                current += perSheetRowCount;
                pageIndex++;
            }

            Save(path, ExcelConstants.DefaultSheetName, pages, options);
        }
        else
        {
            Save(path, ExcelConstants.DefaultSheetName, rows, options);
        }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, IEnumerable<T> rows, ExcelSaveOptions? options = null) => Save(path, ExcelConstants.DefaultSheetName, rows, options);

    public static void Save<T>(string path,
                               string sheet,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        IWorkbook workbook = Path.GetExtension(path).Equals(".xls", StringComparison.CurrentCultureIgnoreCase) ? new HSSFWorkbook() : new XSSFWorkbook();
        try { workbook.Save(path, sheet, rows, itemType, options); }
        finally { workbook.Close(); }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, string sheet, IEnumerable<T> rows, ExcelSaveOptions? options = null)
    {
        IWorkbook workbook = Path.GetExtension(path).Equals(".xls", StringComparison.CurrentCultureIgnoreCase) ? new HSSFWorkbook() : new XSSFWorkbook();
        try { workbook.Save(path, sheet, rows, typeof(T), options); }
        finally { workbook.Close(); }
    }
    public static void Save<T>(string path,
                               IDictionary<string, IEnumerable<T>> sheets,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        IWorkbook workbook = Path.GetExtension(path).Equals(".xls", StringComparison.CurrentCultureIgnoreCase) ? new HSSFWorkbook() : new XSSFWorkbook();
        try { workbook.Save(path, sheets, itemType, options); }
        finally { workbook.Close(); }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, IDictionary<string, IEnumerable<T>> sheets, ExcelSaveOptions? options = null)
    {
        IWorkbook workbook = Path.GetExtension(path).Equals(".xls", StringComparison.CurrentCultureIgnoreCase) ? new HSSFWorkbook() : new XSSFWorkbook();
        try { workbook.Save(path, sheets, typeof(T), options); }
        finally { workbook.Close(); }
    }

    public static void Save<T>(Stream stream,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, rows, itemType, options); }
        finally { workbook.Close(); }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(Stream stream, IEnumerable<T> rows, ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, rows, typeof(T), options); }
        finally { workbook.Close(); }
    }

    public static void Save<T>(Stream stream,
                               string sheet,
                               IEnumerable<T> rows,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, sheet, rows, itemType, options); }
        finally { workbook.Close(); }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(Stream stream, string sheet, IEnumerable<T> rows, ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, sheet, rows, typeof(T), options); }
        finally { workbook.Close(); }
    }

    public static void Save<T>(Stream stream,
                               IDictionary<string, IEnumerable<T>> sheets,
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
                               Type itemType,
                               ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, sheets, itemType, options); }
        finally { workbook.Close(); }
    }

    public static void Save<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(Stream stream, IDictionary<string, IEnumerable<T>> sheets, ExcelSaveOptions? options = null)
    {
        var workbook = new XSSFWorkbook();
        try { workbook.Save<T>(stream, sheets, typeof(T), options); }
        finally { workbook.Close(); }
    }
}
