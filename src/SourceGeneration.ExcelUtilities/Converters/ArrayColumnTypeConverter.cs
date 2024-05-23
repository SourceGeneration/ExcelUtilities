using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class ArrayColumnTypeConverter : ColumnTypeConverter
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.Array;

    public override object? ConvertValue(Type type, object value) => null;

    public override bool Handle(Type type)
    {
        if (!type.IsCompatibleEnumerableGenericInterface(out Type elementType)) return false;
        if (elementType.IsNullableType()) elementType = elementType.GenericTypeArguments[0];

        if (elementType.IsEnum) return true;

        if (elementType == typeof(string) ||
            elementType == typeof(bool) ||
            elementType == typeof(DateTime) ||
            elementType == typeof(DateTimeOffset) ||
            elementType == typeof(TimeSpan) ||
            elementType == typeof(Guid) ||
            elementType == typeof(Uri) ||
            elementType.IsNumberType())
            return true;

#if NET6_0_OR_GREATER
        if(elementType == typeof(DateOnly) || elementType == typeof(TimeOnly)) return true;
#endif

        return false;
    }
}


