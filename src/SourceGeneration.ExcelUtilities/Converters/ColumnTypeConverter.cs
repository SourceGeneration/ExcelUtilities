using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public abstract class ColumnTypeConverter
{
    public abstract bool Handle(Type type);

    public abstract ExcelColumnDataType GetValueType(Type type);

    public abstract object? ConvertValue(Type type, object value);
}

public abstract class ColumnTypeConverter<T> : ColumnTypeConverter
{
    public override bool Handle(Type type) => type.IsAssignableFrom(typeof(T));
}

