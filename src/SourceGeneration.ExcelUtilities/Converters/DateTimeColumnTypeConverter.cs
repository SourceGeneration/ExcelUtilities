using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class DateTimeColumnTypeConverter : ColumnTypeConverter<DateTime>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.DateTime;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is DateTime datetime)
            return datetime;

        if (value is string str && DateTime.TryParse(str, out datetime))
            return datetime;

        return null;
    }
}

