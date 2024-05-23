using System;

namespace SourceGeneration.ExcelUtilities.Converters;

public class NumberColumnTypeConverter : ColumnTypeConverter
{
    public override object? ConvertValue(Type type, object value)
    {
        try
        {
            return Type.GetTypeCode(type) switch
            {
                TypeCode.Double => Convert.ToDouble(value),
                TypeCode.Single => Convert.ToSingle(value),
                TypeCode.Decimal => Convert.ToDecimal(value),
                TypeCode.Byte => Convert.ToByte(value),
                TypeCode.Int16 => Convert.ToInt16(value),
                TypeCode.Int32 => Convert.ToInt32(value),
                TypeCode.Int64 => Convert.ToInt64(value),
                TypeCode.SByte => Convert.ToSByte(value),
                TypeCode.UInt16 => Convert.ToUInt16(value),
                TypeCode.UInt32 => Convert.ToUInt32(value),
                TypeCode.UInt64 => Convert.ToUInt64(value),
                _ => null,
            };
        }
        catch
        {
            return null;
        }
    }

    public override ExcelColumnDataType GetValueType(Type type)
    {
        if (type == typeof(long) || type == typeof(ulong)) return ExcelColumnDataType.String;
        if (type.IsIntegerType()) return ExcelColumnDataType.Integer;
        else return ExcelColumnDataType.Number;
    }

    public override bool Handle(Type type)
    {
        if (type.IsNullableType())
            type = type.GenericTypeArguments[0];

        return type.IsNumberType();
    }
}