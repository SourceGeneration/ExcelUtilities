using System.Runtime.CompilerServices;

namespace SourceGeneration.ExcelUtilities;

internal static class StringExtensions
{
    public static readonly char[] BlankChars =
    [
        '\u0020',
        '\u00A0',

        '\u2001',
        '\u2002',
        '\u2003',
        '\u2004',
        '\u2005',
        '\u2006',
        '\u2007',
        '\u2008',
        '\u2009',
        '\u200A',
        '\u200B',
        '\u200C',
        '\u200D',

        '\u2060',

        '\u202C',
        '\u202D',
        '\u202E',

        '\u205F',

        '\u3000',
        '\uFEFF',
    ];

    static readonly char[] AllSpaceChars =
    [
        '\t',//tab
        '\r',
        '\n',

        '\u0020',
        '\u00A0',

        '\u2001',
        '\u2002',
        '\u2003',
        '\u2004',
        '\u2005',
        '\u2006',
        '\u2007',
        '\u2008',
        '\u2009',
        '\u200A',
        '\u200B',
        '\u200C',
        '\u200D',

        '\u2060',

        '\u202C',
        '\u202D',
        '\u202E',

        '\u205F',

        '\u3000',
        '\uFEFF',
    ];

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string? TrimSpace(this string? s) => s?.Trim(AllSpaceChars);
}