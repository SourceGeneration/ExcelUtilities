#if NET6_0_OR_GREATER

using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class TimeOnlyColumnTypeConverter : ColumnTypeConverter<TimeOnly>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Time;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is TimeOnly time)
            return time;

        if (value is DateTime datetime)
            return TimeOnly.FromDateTime(datetime);

        if (value is string str && TimeOnly.TryParse(str, out time))
            return time;

        return null;
    }
}

#endif