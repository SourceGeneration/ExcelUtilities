using NPOI.SS.UserModel;
using System;

namespace SourceGeneration.ExcelUtilities;

#pragma warning disable IDE0038 // 使用模式匹配

internal static class ICellExtensions
{
    public static void SetCallNumberValue(this ICell cell, object value)
    {
        if (value is byte) cell.SetCellValue((byte)value);
        else if (value is sbyte) cell.SetCellValue((sbyte)value);
        else if (value is short) cell.SetCellValue((short)value);
        else if (value is ushort) cell.SetCellValue((ushort)value);
        else if (value is int) cell.SetCellValue((int)value);
        else if (value is uint) cell.SetCellValue((uint)value);
        //else if (value is long) cell.SetCellValue(value.ToString());
        //else if (value is ulong) cell.SetCellValue(value.ToString());
        else if (value is long) cell.SetCellValue((long)value);
        else if (value is ulong) cell.SetCellValue((ulong)value);
        else if (value is char) cell.SetCellValue((char)value);
        else if (value is float) cell.SetCellValue(Convert.ToDouble((float)value));
        else if (value is double) cell.SetCellValue((double)value);
        else if (value is decimal) cell.SetCellValue(decimal.ToDouble((decimal)value));
    }

    public static void SetCellValue(this ICell cell, object value)
    {
        if (value == null)
            return;

        if (value is bool) cell.SetCellValue((bool)value);
        else if (value is string) cell.SetCellValue((string)value);
        else if (value is char) cell.SetCellValue(((char)value).ToString());
        else if (value is byte) cell.SetCellValue((byte)value);
        else if (value is sbyte) cell.SetCellValue((sbyte)value);
        else if (value is short) cell.SetCellValue((short)value);
        else if (value is ushort) cell.SetCellValue((ushort)value);
        else if (value is int) cell.SetCellValue((int)value);
        else if (value is uint) cell.SetCellValue((uint)value);
        else if (value is long) cell.SetCellValue((long)value);
        else if (value is ulong) cell.SetCellValue((ulong)value);
        else if (value is char) cell.SetCellValue((char)value);
        else if (value is float) cell.SetCellValue(Convert.ToDouble((float)value));
        else if (value is double) cell.SetCellValue((double)value);
        else if (value is decimal) cell.SetCellValue(decimal.ToDouble((decimal)value));
        else if (value is DateTime) cell.SetCellValue((DateTime)value);
        else if (value is DateTimeOffset) cell.SetCellValue(((DateTimeOffset)value).DateTime);
        else if (value is Uri) cell.SetCellValue(((Uri)value).ToString());
    }

    public static bool TryGetDateTimeValue(this ICell cell, out DateTime? datetime)
    {
        if (cell.CellType == CellType.Numeric)
        {
            try
            {
                datetime = cell.DateCellValue;
                return true;
            }
            catch { }
        }
        else if (cell.CellType == CellType.String)
        {
            if (DateTime.TryParse(cell.StringCellValue, out DateTime dt))
            {
                datetime = dt;
                return true;
            }
        }

        datetime = DateTime.MinValue;
        return false;
    }

    public static bool TryGetBooleanValue(this ICell cell, out bool value)
    {
        value = false;
        var stringValue = cell.GetStringValue()?.TrimSpace();

        if (string.IsNullOrEmpty(stringValue))
            return false;

        if (bool.TryParse(stringValue, out value))
            return true;

        if (stringValue == "是")
        {
            value = true;
            return true;
        }

        if (stringValue == "否")
        {
            value = false;
            return true;
        }

        return false;
    }

    public static string GetStringValue(this ICell cell)
    {
        if (cell.CellType == CellType.Formula)
            return cell.GetStringValue(cell.CachedFormulaResultType);
        else
            return cell.GetStringValue(cell.CellType);
    }

    public static bool TryGetNumericValue(this ICell cell, out double value)
    {
        if (cell.CellType == CellType.Formula && cell.TryGetNumericValue(cell.CachedFormulaResultType, out value))
        {
            return true;
        }
        else if (cell.TryGetNumericValue(cell.CellType, out value))
        {
            return true;
        }
        return false;
    }


#if NET8_0_OR_GREATER

    public static bool TryGetDateOnlyValue(this ICell cell, out DateOnly? date)
    {
        if (cell.CellType == CellType.Numeric)
        {
            try
            {
                if (cell.DateOnlyCellValue.HasValue)
                {
                    date = cell.DateOnlyCellValue.Value;
                }
                else if (cell.DateCellValue.HasValue)
                {
                    date = DateOnly.FromDateTime(cell.DateCellValue.Value);
                }
                else
                {
                    date = null;
                }
                return true;
            }
            catch { }
        }
        else if (cell.CellType == CellType.String)
        {
            if (DateOnly.TryParse(cell.StringCellValue, out DateOnly d))
            {
                date = d;
                return true;
            }
        }

        date = DateOnly.MinValue;
        return false;
    }


    public static bool TryGetTimeOnlyValue(this ICell cell, out TimeOnly? time)
    {
        if (cell.CellType == CellType.Numeric)
        {
            try
            {
                if (cell.TimeOnlyCellValue.HasValue)
                {
                    time = cell.TimeOnlyCellValue.Value;
                }
                else if (cell.DateCellValue.HasValue)
                {
                    time = TimeOnly.FromDateTime(cell.DateCellValue.Value);
                }
                else
                {
                    time = null;
                }
                return true;
            }
            catch { }
        }
        else if (cell.CellType == CellType.String)
        {
            if (TimeOnly.TryParse(cell.StringCellValue, out TimeOnly t))
            {
                time = t;
                return true;
            }
        }

        time = TimeOnly.MinValue;
        return false;
    }

#endif

    private static string GetStringValue(this ICell cell, CellType type)
    {

        switch (type)
        {
            case CellType.String: return cell.StringCellValue;
            case CellType.Boolean: return cell.BooleanCellValue.ToString();
            case CellType.Numeric:
                {
                    var format = cell.CellStyle.GetDataFormatString();
                    if (format == "m/d/yy")
                    {
                        return new DateTime(1900, 1, 1).AddDays(cell.NumericCellValue - 2).ToString("yyyy-MM-dd");
                    }
                    return cell.NumericCellValue.ToString();
                }
            default: return string.Empty;
        }
    }

    private static bool TryGetNumericValue(this ICell cell, CellType type, out double value)
    {
        if (type == CellType.Numeric)
        {
            value = cell.NumericCellValue;
            return true;
        }

        if (type == CellType.String && double.TryParse(cell.StringCellValue, out value))
        {
            return true;
        }

        value = default;
        return false;
    }

    public static bool TryGetValue(this ICell cell, ExcelColumnDataType valueType, out object? value)
    {
        switch (valueType)
        {
            case ExcelColumnDataType.Boolean:
                if (cell.TryGetBooleanValue(out bool boolean))
                {
                    value = boolean;
                    return true;
                }
                var stringValue = cell.GetStringValue();
                if (stringValue == "是")
                {
                    value = true;
                    return true;
                }
                else if (stringValue == "否")
                {
                    value = false;
                    return true;
                }
                break;
            case ExcelColumnDataType.Integer:
            case ExcelColumnDataType.Number:
                if (cell.TryGetNumericValue(out double number))
                {
                    value = number;
                    return true;
                }
                break;
            case ExcelColumnDataType.DateTime:
                if (cell.TryGetDateTimeValue(out DateTime? datetime))
                {
                    value = datetime;
                    return true;
                }
                break;
            case ExcelColumnDataType.Date:
#if NET8_0_OR_GREATER
                if (cell.TryGetDateOnlyValue(out DateOnly? date))
                {
                    value = date;
                    return true;
                }
#else
                if (cell.TryGetDateTimeValue(out DateTime? date))
                {
                    value = date;
                    return true;
                }
#endif
                break;
            case ExcelColumnDataType.Time:
#if NET8_0_OR_GREATER
                if (cell.TryGetTimeOnlyValue(out TimeOnly? time))
                {
                    value = time;
                    return true;
                }
#else
                if (cell.TryGetDateTimeValue(out DateTime? time))
                {
                    value = time;
                    return true;
                }
#endif
                break;
            case ExcelColumnDataType.Enum:
            case ExcelColumnDataType.String:
                value = cell.GetStringValue();
                return true;
            default: break;
        }

        value = null;
        return false;

    }

}
