using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public sealed class UnknownColumnTypeConverter : ColumnTypeConverter
{
    public override bool Handle(Type type) => true;
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Unknown;
    public override object? ConvertValue(Type type, object value) => null;
}

