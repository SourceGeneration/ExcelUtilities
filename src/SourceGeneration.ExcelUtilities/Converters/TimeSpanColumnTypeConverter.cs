using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class TimeSpanColumnTypeConverter : ColumnTypeConverter<TimeSpan>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.String;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is TimeSpan timespan)
            return timespan;

        if (value is string str && TimeSpan.TryParse(str, out timespan))
            return timespan;

        return null;
    }
}

