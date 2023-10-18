using System;

namespace DocxSteganography.Core;

public class Result
{
    public string? Error { get; }
    public int ExitCode { get; }
    public bool IsSuccess => Error == null;

    public Result() : this(null, 0)
    {
    }
    
    public Result(string? error, int exitCode)
    {
        Error = error;
        ExitCode = exitCode;
    }

    public static implicit operator Result(int code)
    {
        if (code != 0)
            throw new ArgumentException("Exit code must 0 for implicit conversion!");
        
        return new Result();
    }
    
    public static implicit operator int(Result result)
    {
        return result.ExitCode;
    }

    public Result OnSuccess(Action<Result> handler)
    {
        if (IsSuccess)
            handler(this);

        return this;
    }

    public Result OnError(Action<Result> handler)
    {
        if (!IsSuccess)
            handler(this);

        return this;
    }
}