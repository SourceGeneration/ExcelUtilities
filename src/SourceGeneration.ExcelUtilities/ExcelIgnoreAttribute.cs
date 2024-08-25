using System;

namespace SourceGeneration.ExcelUtilities;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field,AllowMultiple = false,Inherited = true)]
public sealed class ExcelIgnoreAttribute : Attribute { }
