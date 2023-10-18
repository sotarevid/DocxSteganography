using CommandLine;

namespace DocxSteganography.CLI;

public class Options
{
    [Option('i', "input", Required = true, HelpText = "The file to read secret message from.")]
    public string PathToTxt { get; init; }

    [Option('o', "output", Required = true, HelpText = "The file to hide the secret message in.")]
    public string PathToDocx { get; init; }
    
    [Option("debug")]
    public bool Debug { get; init; }
}