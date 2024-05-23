using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class BooleanColumnTypeConverter : ColumnTypeConverter<bool>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Boolean;

    public override object? ConvertValue(Type type, object value)
    {
        if (value == null)
            return null;

        if (value is bool boolean)
            return boolean;

        if (value is string stringValue)
        {
            if (bool.TryParse(stringValue, out boolean))
                return boolean;

            if (stringValue == "是") return true;
            if (stringValue == "否") return false;
        }

        return null;
    }
}

