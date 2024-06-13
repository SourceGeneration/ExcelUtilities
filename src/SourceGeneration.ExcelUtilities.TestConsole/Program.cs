using MathNet.Numerics.Random;
using SourceGeneration.ExcelUtilities;
using SourceGeneration.Reflection;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Reflection;


TestDataModel[] models = new TestDataModel[10];

for (int i = 0; i < 10; i++)
{
    models[i] = new TestDataModel
    {
        BoolValue = Random.Shared.NextBoolean(),
        Byte = (byte)Random.Shared.Next(byte.MinValue, byte.MaxValue),
        SByte = (sbyte)Random.Shared.Next(sbyte.MinValue, sbyte.MaxValue),
        Int16 = (short)Random.Shared.Next(short.MinValue, short.MaxValue),
        UInt16 = (ushort)Random.Shared.Next(ushort.MinValue, ushort.MaxValue),
        Int32 = Random.Shared.Next(),
        UInt32 = (uint)Random.Shared.NextInt64(uint.MinValue, uint.MaxValue),
        Int64 = Random.Shared.NextInt64(),
        UInt64 = (ulong)Random.Shared.NextInt64(0, long.MaxValue),
        Float = Random.Shared.NextSingle(),
        Double = Random.Shared.NextDouble(),
        Decimal = Random.Shared.NextDecimal(),
        DateTime = DateTime.Now,
        DateTimeOffset = DateTimeOffset.Now,
        Time = TimeOnly.FromDateTime(DateTime.Now),
        TimeSpan = TimeSpan.FromMilliseconds(Random.Shared.Next()),
        Date = DateOnly.FromDateTime(DateTime.Now),
        Color = Color.FromArgb(Random.Shared.Next(0, int.MaxValue)),
        String = "abcdefg",
        EnumValue = Random.Shared.NextBoolean() ? TestEnum.A : TestEnum.B,
        Uri = new Uri("https://www.microsoft.com"),
        IntegerArray = [null, 1, 2, 3, 4, 5, 6, 7, 8, 9],
        DateTimeArray = [DateTime.Now, DateTime.Now.AddDays(1), DateTime.Now.AddDays(2)],
        StringArray = ["abc", "efg", "xyz"],
        Timestamp1 = DateTimeOffset.UtcNow.ToUnixTimeSeconds(),
        Timestamp2 = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
    };
}

var options = new ExcelSaveOptions<TestDataModel>
{
    AutoSizeRow = true,
};
options.Column(x => x.BoolValue, "B", "Yes:No");
options.Column(x => x.Timestamp1, "C", null, true);

ExcelUtility.Save("test-clr2.xlsx", models, options);

ExcelUtility.Save("test-clr.xlsx", models, new ExcelSaveOptions
{
    AutoSizeRow = true,
});

[SourceReflection]
public class TestDataModel : IComparable
{
    public bool BoolValue { get; set; }

    public byte Byte { get; set; }
    public sbyte SByte { get; set; }
    public short Int16 { get; set; }
    public ushort UInt16 { get; set; }
    public int Int32 { get; set; }
    public uint UInt32 { get; set; }
    public long Int64 { get; set; }
    public ulong UInt64 { get; set; }
    public decimal Decimal { get; set; }

    [Display(Name = "单精度浮点数")] public float Float { get; set; }

    [Display(Name = "双精度浮点数")] public double Double { get; set; }

    [Display(Name = "日期时间")] public DateTime DateTime { get; set; }

    [Display(Name = "日期时间2")] public DateTimeOffset DateTimeOffset { get; set; }

    [Display(Name = "日期")] public DateOnly Date { get; set; }

    [Display(Name = "时间")] public TimeOnly Time { get; set; }

    [Display(Name = "时间间隔")] public TimeSpan TimeSpan { get; set; }

    [Display(Name = "颜色")] public Color Color { get; set; }

    [Display(Name = "连接")] public Uri Uri { get; set; }

    [Timestamp] public long Timestamp1 { get; set; }
    [DataType(DataType.DateTime)] public long Timestamp2 { get; set; }
    public TestEnum EnumValue { get; set; }

    public string String { get; set; }
    public string[] StringArray { get; set; }

    public int?[] IntegerArray { get; set; }

    public DateTime?[] DateTimeArray { get; set; }

    public int CompareTo(object? obj)
    {
        throw new NotImplementedException();
    }
}

[SourceReflection]
public enum TestEnum
{
    [Description("选项A")] A,
    [Description("选项B")] B,
}
