using SourceGeneration.ExcelUtilities.Converters;
using SourceGeneration.Reflection;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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

    private Dictionary<Type, ExcelColumnBase[]> _columns = [];

    internal virtual IReadOnlyList<ExcelColumnBase> GetColumns(
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
        Type type)
    {
        if (_columns.TryGetValue(type, out var columns))
        {
            return columns;
        }

        var typeInfo = SourceReflector.GetType(type, true);
        var members = typeInfo?.GetFieldsAndProperties()?.Where(x =>
            x.IsStatic == false &&
            x.MemberInfo.GetCustomAttribute<ExcelIgnoreAttribute>() == null &&
            (x.Accessibility == SourceAccessibility.Public || 
            x.Accessibility == SourceAccessibility.Internal || 
            x.Accessibility == SourceAccessibility.ProtectedOrInternal))?.ToList() ?? [];

        columns = new ExcelColumnBase[members.Count];
        for (int i = 0; i < members.Count; i++)
        {
            var memberInfo = members[i];
            var display = memberInfo.MemberInfo.GetCustomAttribute<DisplayAttribute>();
            var columnOptions = new ExcelColumnOptions
            {
                Title = display?.GetName() ?? memberInfo.Name,
                Order = display?.GetOrder() ?? 0,
                MemberInfo = memberInfo,
                IsTimestamp = memberInfo.MemberInfo.GetCustomAttribute<DataTypeAttribute>()?.DataType == DataType.DateTime,
                Format = memberInfo.MemberInfo.GetCustomAttribute<DisplayFormatAttribute>()?.DataFormatString,
            };

            columns[i] = new ExcelColumn(this, columnOptions);
        }

        columns = [.. columns.OrderBy(x => x.Order)];
        _columns.Add(type, columns);

        return columns;
    }
}
