using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace SourceGeneration.ExcelUtilities;

public static partial class ExcelUtility
{
    public static IEnumerable<T> Load<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(Stream stream, int sheetNum, ExcelLoadOptions? options = null) where T : class, new()
    {
        XSSFWorkbook workbook = new(stream);
        return workbook.GetSheetAt(sheetNum).ReadRows<T>(options);
    }

    public static IEnumerable<T> Load<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(Stream stream, ExcelLoadOptions? options = null) where T : class, new()
    {
        XSSFWorkbook workbook = new(stream);
        return workbook.GetSheetAt(0).ReadRows<T>(options);
    }

    public static IEnumerable<T> Load<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, int sheetNum, ExcelLoadOptions? options = null) where T : class, new()
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found", path);

        XSSFWorkbook workbook = new(path);
        return workbook.GetSheetAt(sheetNum).ReadRows<T>(options);
    }

    public static IEnumerable<T> Load<
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, string sheetName, ExcelLoadOptions? options = null) where T : class, new()
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found", path);

        XSSFWorkbook workbook = new(path);
        return workbook.GetSheet(sheetName).ReadRows<T>(options);
    }

    public static IEnumerable<T> Load<
#if NET5_0_OR_GREATER 
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    T>(string path, ExcelLoadOptions? options = null) where T : class, new()
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("File not found", path);
        XSSFWorkbook workbook = new(path);
        return workbook.GetSheetAt(0).ReadRows<T>(options);
    }
}
