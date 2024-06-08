using NPOI.SS.Formula.Functions;
using SourceGeneration.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Threading;
using static NPOI.HSSF.UserModel.HeaderFooter;

namespace SourceGeneration.ExcelUtilities;

/// <inheritdoc/>
public class ExcelSaveOptions : ExcelOptions
{
    public static readonly ExcelSaveOptions Default = new();

    public bool AutoCreateHeader { get; set; } = true;

    public bool AutoCreateRowNumberColumn { get; set; } = true;

    public bool AutoSizeColumn { get; set; } = true;

    public bool AutoSizeRow { get; set; } = false;

    public int RowHeight { get; set; } = 20;

    public string ArrayItemPrefix { get; set; } = "· ";

    public string ArrayItemSeparator { get; set; } = "\n";

    internal override IReadOnlyList<ExcelColumnBase> GetColumns(
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
        Type type)
    {
        var columns = base.GetColumns(type);

        if (!AutoCreateRowNumberColumn)
        {
            return columns;
        }

        var culture = Thread.CurrentThread.CurrentUICulture.Name;
        var title = (culture == "zh" || culture.StartsWith("zh-")) ? "序号" : "No.";

        return [new ExcelRowNumberColumn(title), .. columns];
    }
}

public class ExcelSaveOptions<
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
T> : ExcelSaveOptions
{
    private readonly List<ExcelColumn> _columns = [];
    private SourceTypeInfo? _typeInfo;

    public void Column(Expression<Func<T, object>> selector, string? title = null, string? format = null, bool timestamp = false)
    {
        string? memberName;

        var exp = selector.Body;
        if (exp is UnaryExpression unary && unary.NodeType == ExpressionType.Convert)
        {
            exp = unary.Operand;
        }

        if (exp is MemberExpression member)
        {
            memberName = member.Member.Name;
        }
        else
        {
            throw new InvalidCastException("不支持的表达式");
        }

        _typeInfo ??= SourceReflector.GetRequiredType<T>(true);
        var memberInfo = _typeInfo.GetFieldOrProperty(memberName);

        _columns.Add(new ExcelColumn(this, new ExcelColumnOptions
        {
            MemberInfo = memberInfo!,
            Title = title ?? memberInfo!.Name,
            Format = format,
            IsTimestamp = timestamp,
        }));
    }

    internal override IReadOnlyList<ExcelColumnBase> GetColumns(
#if NET5_0_OR_GREATER
    [DynamicallyAccessedMembers(Dynamically.DefaultAccessMembers)]
#endif
    Type type)
    {
        if (type == typeof(T))
        {
            var culture = Thread.CurrentThread.CurrentUICulture.Name;
            var title = (culture == "zh" || culture.StartsWith("zh-")) ? "序号" : "No.";

            return [new ExcelRowNumberColumn(title), .. _columns];
        }
        return base.GetColumns(type);
    }
}

