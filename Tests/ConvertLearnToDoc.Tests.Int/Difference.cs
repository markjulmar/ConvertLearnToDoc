using System;

namespace ConvertLearnToDoc.Tests.Int;

public class Difference<T> where T: class
{
    public int? LineNumber { get; init; }
    public string? Filename1 { get; init; }
    public string? Filename2 { get; init; }
    public T? Value1 { get; init; }
    public T? Value2 { get; init; }

    public override string ToString()
    {
        return $"{LineNumber}:{Environment.NewLine}"
               + $"#1: {Value1?.ToString() ?? "(missing)"}{Environment.NewLine}"
               + $"#2: {Value2?.ToString() ?? "(missing)"}{Environment.NewLine}";
    }
}