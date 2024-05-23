using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

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
        var name = (culture == "zh" || culture.StartsWith("zh-")) ? "序号" : "No.";

        return [new ExcelRowNumberColumn(name), .. columns];
    }
}