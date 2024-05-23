#if NET6_0_OR_GREATER

using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class DateOnlyColumnTypeConverter : ColumnTypeConverter<DateOnly>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Date;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is DateOnly time)
            return time;

        if (value is DateTime datetime)
            return DateOnly.FromDateTime(datetime);

        if (value is string str && DateOnly.TryParse(str, out time))
            return time;

        return null;
    }
}

#endif