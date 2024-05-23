using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class EnumColumnTypeConverter : ColumnTypeConverter
{
    public override bool Handle(Type type) => type.IsEnum;
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Enum;

    public override object? ConvertValue(Type type, object value)
    {
#if NET6_0_OR_GREATER
        if (value is string str && Enum.TryParse(type, str, true, out object? @enum))
            return @enum;
#else
        if (value is string str)
        {
            try
            {
                return Enum.Parse(type, str, true);
            }
            catch { }
        }
#endif

        try
        {
            return Convert.ChangeType(value, type);
        }
        catch { }
        return null;
    }
}