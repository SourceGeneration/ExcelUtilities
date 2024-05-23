using SourceGeneration.ExcelUtilities.Converters;
using SourceGeneration.Reflection;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;

namespace SourceGeneration.ExcelUtilities;

public class ExcelOptions
{
    private readonly static ColumnTypeConverter UnknownTypeConverter = new UnknownColumnTypeConverter();

    private readonly static List<ColumnTypeConverter> DefaultConverters =
    [
            new BooleanColumnTypeConverter(),
            new EnumColumnTypeConverter(),
            new NumberColumnTypeConverter(),
            new StringColumnTypeConverter(),
            new DateTimeColumnTypeConverter(),
            new DateTimeOffsetColumnTypeConverter(),
            new TimeSpanColumnTypeConverter(),
            new GuidColumnTypeConverter(),
            new UriColumnTypeConverter(),
            new ArrayColumnTypeConverter(),
    #if NET6_0_OR_GREATER
            new DateOnlyColumnTypeConverter(),
            new TimeOnlyColumnTypeConverter(),
    #endif
        ];

    public List<ColumnTypeConverter> Converters { get; } = [];

    internal ColumnTypeConverter GetConverter(Type type)
    {
        foreach (var converter in Converters)
        {
            if (converter.Handle(type))
                return converter;
        }

        foreach (var converter in DefaultConverters)
        {
            if (converter.Handle(type))
                return converter;
        }
        return UnknownTypeConverter;
    }

    internal virtual IReadOnlyList<ExcelColumnBase> GetColumns(
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
        Type type)
    {
        var typeInfo = SourceReflector.GetType(type, true);
        var members = typeInfo?.GetFieldsAndProperties()?.Where(x =>
            x.IsStatic == false &&
            x.MemberInfo.GetCustomAttribute<NotMappedAttribute>() == null &&
            x.Accessibility == SourceAccessibility.Public)?.ToList() ?? [];

        ExcelColumnBase[] columns = new ExcelColumnBase[members.Count];
        for (int i = 0; i < members.Count; i++)
        {
            columns[i] = new ExcelColumn(this, members[i]);
        }

        return [.. columns.OrderBy(x => x.Order)];
    }
}
