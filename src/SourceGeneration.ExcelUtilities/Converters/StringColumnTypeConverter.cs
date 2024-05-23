using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class StringColumnTypeConverter : ColumnTypeConverter<string>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.String;
    public override object? ConvertValue(Type type, object value) => value?.ToString();
}

