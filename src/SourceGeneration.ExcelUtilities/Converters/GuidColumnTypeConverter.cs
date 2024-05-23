using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class GuidColumnTypeConverter : ColumnTypeConverter<DateTime>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.String;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is Guid guid)
            return guid;

        if (value is string str && Guid.TryParse(str, out guid))
            return guid;

        return null;
    }
}

