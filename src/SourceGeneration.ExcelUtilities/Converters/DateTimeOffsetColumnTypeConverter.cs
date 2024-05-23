using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class DateTimeOffsetColumnTypeConverter : ColumnTypeConverter<DateTimeOffset>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.DateTime;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is DateTimeOffset datetimeoffset)
            return datetimeoffset.DateTime;

        if (value is DateTime datetime)
            return new DateTimeOffset(datetime);

        if (value is string str && DateTimeOffset.TryParse(str, out datetimeoffset))
            return datetimeoffset;

        return null;
    }
}

