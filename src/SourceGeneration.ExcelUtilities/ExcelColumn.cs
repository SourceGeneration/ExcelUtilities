﻿using EnumsNET;
using NPOI.SS.UserModel;
using SourceGeneration.ExcelUtilities.Converters;
using SourceGeneration.Reflection;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SourceGeneration.ExcelUtilities;

public class ExcelColumnOptions
{
    public ISourceFieldOrPropertyInfo MemberInfo { get; set; } = null!;
    public string? Title { get; set; }
    public int Order { get; set; }
    public string? Format { get; set; }
    public bool IsTimestamp { get; set; }
}

public class ExcelColumn : ExcelColumnBase
{
    private readonly ColumnTypeConverter _converter;
    private readonly ISourceFieldOrPropertyInfo _memberInfo;
    private readonly Type _valueType;
    private readonly bool _timestamp;
    private readonly string? _format;
    private readonly Dictionary<object, string> _enumMembers = [];
    private readonly string? _trueValue;
    private readonly string? _falseValue;

    public ExcelColumn(ExcelOptions options, ExcelColumnOptions columnOptions)
    {
        _memberInfo = columnOptions.MemberInfo;

        _timestamp = columnOptions.IsTimestamp;

        Title = columnOptions.Title;
        Order = columnOptions.Order;


        if (_memberInfo.MemberType.IsNullableType())
            _valueType = _memberInfo.MemberType.GenericTypeArguments[0];
        else
            _valueType = _memberInfo.MemberType;

        _converter = options.GetConverter(_valueType);

        if (_timestamp)
        {
            DataType = ExcelColumnDataType.Integer;
        }
        else
        {
            DataType = _converter.GetValueType(_valueType);
        }

        if (DataType == ExcelColumnDataType.Boolean)
        {
            _format = columnOptions.Format;
            if (_format != null)
            {
                var index = _format.IndexOf(':');
                if (index >= 0)
                {
                    _trueValue = _format.Substring(0, index);
                    if (index < _format.Length - 1)
                        _falseValue = _format.Substring(index + 1);
                }
            }
        }
        else if (DataType == ExcelColumnDataType.Enum)
        {
            foreach (var member in SourceReflector.GetRequiredType(_valueType, true).DeclaredFields.Where(x => x.IsStatic))
            {
                var key = member.GetValue(null)!;
                var value =
                    member.FieldInfo.GetCustomAttribute<DisplayAttribute>()?.Name ??
                    member.FieldInfo.GetCustomAttribute<DescriptionAttribute>()?.Description ??
                    key.ToString();
                _enumMembers.Add(Enum.ToObject(_valueType, key!), value!);
            }
        }
    }

    public override void ReadCell(ICell cell, object obj)
    {
        if (cell == null)
            return;

        if (!cell.TryGetValue(DataType, out object? value))
            return;

        if (value == null)
            return;

        value = _converter.ConvertValue(_valueType, value);
        if (value == null)
            return;

        if (!_valueType.IsAssignableFrom(value.GetType()))
            return;

        _memberInfo.SetValue(obj, value);
    }

    public override void WriteCell(ICell cell, object? obj, int rowNumber, ExcelRowStyles styles, ExcelSaveOptions options)
    {
        if (obj == null)
        {
            cell.CellStyle = styles.LeftCellStyle;
            return;
        }

        var value = GetColumnValue(obj, options);
        if (value == null)
        {
            cell.CellStyle = styles.LeftCellStyle;
            return;
        }

        switch (DataType)
        {
            case ExcelColumnDataType.Boolean:
                WriteBooleanValue(cell, value, styles);
                break;
            case ExcelColumnDataType.String:
                WriteStringValue(cell, value, styles);
                break;
            case ExcelColumnDataType.Integer:
                if (_timestamp)
                {
                    var timestamp = (long)value;
                    if (timestamp > 253402300799)
                        WriteDateTimeValue(cell, DateTimeOffset.FromUnixTimeMilliseconds(timestamp).LocalDateTime, styles);
                    else
                        WriteDateTimeValue(cell, DateTimeOffset.FromUnixTimeSeconds(timestamp).LocalDateTime, styles);
                }
                else
                {
                    WriteIntegerValue(cell, value, styles);
                }
                break;
            case ExcelColumnDataType.Number:
                WriteNumberValue(cell, value, styles);
                break;
            case ExcelColumnDataType.DateTime:
                WriteDateTimeValue(cell, value, styles);
                break;
            case ExcelColumnDataType.Enum:
                WriteEnumValue(cell, value, styles);
                break;
            case ExcelColumnDataType.Date:
#if NET6_0_OR_GREATER
                WriteDateValue(cell, value, styles);
#else
                WriteDateTimeValue(cell, value, styles);
#endif
                break;
            case ExcelColumnDataType.Time:
#if NET6_0_OR_GREATER
                WriteTimeValue(cell, value, styles);
#else
                WriteDateTimeValue(cell, value, styles);
#endif
                break;
            default:
                WriteDefaultValue(cell, value, styles);
                break;
        }
    }

    protected virtual object? GetColumnValue(object obj, ExcelSaveOptions options)
    {
        object? value = _memberInfo.GetValue(obj);

        if (DataType == ExcelColumnDataType.Array)
        {
            if (value is IEnumerable enumerable)
            {
                var array = new List<string>();
                foreach (var item in enumerable) array.Add(options.ArrayItemPrefix + item?.ToString());
                return string.Join(options.ArrayItemSeparator, array);
            }
        }

        return value;
    }

    protected virtual void WriteBooleanValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.String);
        cell.CellStyle = styles.CenterCellStyle;

        if (_format != null)
        {
            cell.SetCellValue((bool)value ? _trueValue : _falseValue);
        }
        else
        {
            if (System.Globalization.CultureInfo.CurrentUICulture.Name == "zh-CN")
                cell.SetCellValue((bool)value ? "是" : "否");
            else
                cell.SetCellValue(value);
        }
    }

    protected virtual void WriteStringValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.String);
        cell.CellStyle = styles.TextCellStyle;
        cell.CellStyle.WrapText = true;
        if (value is string v)
            cell.SetCellValue(v);
        else
            cell.SetCellValue(value.ToString());
    }

    protected virtual void WriteIntegerValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.Numeric);
        cell.CellStyle = styles.IntegerCellStyle;
        cell.SetCallNumberValue(value);
    }

    protected virtual void WriteNumberValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.Numeric);
        cell.CellStyle = styles.NumberCellStyle;
        cell.SetCallNumberValue(value);
    }

    protected virtual void WriteEnumValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.String);
        cell.CellStyle = styles.TextCellStyle;
        _enumMembers.TryGetValue(value, out string? cellValue);
        cell.SetCellValue(cellValue ?? value.ToString());
    }

    protected virtual void WriteDateTimeValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.Numeric);
        cell.CellStyle = styles.DateTimeCellStyle;

        if (value is DateTime time) cell.SetCellValue(time);
        else if (value is DateTimeOffset offset) cell.SetCellValue(offset.DateTime);
    }

#if NET6_0_OR_GREATER
    protected virtual void WriteDateValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.Numeric);
        cell.CellStyle = styles.DateCellStyle;
        cell.SetCellValue(((DateOnly)value).ToDateTime(TimeOnly.MinValue));
    }

    protected virtual void WriteTimeValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.String);
        cell.CellStyle = styles.TimeCellStyle;
        cell.SetCellValue(((TimeOnly)value).ToString("HH:mm:ss"));
    }

#endif

    protected virtual void WriteDefaultValue(ICell cell, object value, ExcelRowStyles styles)
    {
        cell.SetCellType(CellType.String);
        cell.CellStyle = styles.TextCellStyle;
        cell.SetCellValue(value.ToString());
    }

    //public override int ComputeColumnSize()
    //{
    //    switch (ValueType)
    //    {
    //        //case NpoiColumnValueType.Boolean:
    //        //    return 256 * 6;
    //        //case NpoiColumnValueType.Integer:
    //        //case NpoiColumnValueType.Number:
    //        //    return 256 * 12;
    //        case NpoiColumnValueType.Date:
    //        case NpoiColumnValueType.Time:
    //            return 256 * 12;
    //        default: return 0;
    //    }
    //}
}
