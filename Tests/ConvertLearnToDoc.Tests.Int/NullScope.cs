using System;

namespace ConvertLearnToDoc.Tests.Int;

public class NullScope : IDisposable
{
    public static NullScope Instance { get; } = new NullScope();

    private NullScope() { }

    public void Dispose() { }
}
