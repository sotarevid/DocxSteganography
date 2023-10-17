using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using CommandLine;
using DocumentFormat.OpenXml.Packaging;
using DocxStenography.Core;

namespace DocxStenography.CLI;

public static class Program
{
    public static int Main(string[] args)
    {
        return Parser.Default.ParseArguments<Options>(args)
            .MapResult(OnParseSuccess, OnParseFail);
    }

    private static int OnParseSuccess(Options options)
    {
        var message = new Regex(@"\s+", RegexOptions.Compiled)
            .Replace(File.ReadAllText(options.PathToTxt), "");
        
        using var document = WordprocessingDocument.Open(options.PathToDocx, true);
        
        return new Stenographer(options.Debug)
            .HideMessage(message, document)
            .OnError(r => Console.WriteLine(r.Error));
    }
    
    private static int OnParseFail(IEnumerable<Error> errors)
    {
        foreach (var error in errors)
            Console.WriteLine(error);

        return 1;
    }
}