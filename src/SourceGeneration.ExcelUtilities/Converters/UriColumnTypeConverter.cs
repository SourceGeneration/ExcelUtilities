using System;
using System.Collections.Generic;

namespace SourceGeneration.ExcelUtilities.Converters;

public class UriColumnTypeConverter : ColumnTypeConverter<Uri>
{
    public override ExcelColumnDataType GetValueType(Type type) => ExcelColumnDataType.String;

    public override object? ConvertValue(Type type, object value)
    {
        if (value is Uri uri)
            return uri;

        if (value is string str && Uri.IsWellFormedUriString(str, UriKind.RelativeOrAbsolute))
            return new Uri(str);

        return null;
    }
}