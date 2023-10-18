using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxSteganography.Core;

public class Steganographer
{
    private readonly int _fontSizeModifier;

    public Steganographer(bool debugMode = false)
    {
        _fontSizeModifier = debugMode ? 100 : -1;
    }
    
    public Result HideMessage(string message, WordprocessingDocument document)
    {
        if (document.MainDocumentPart?.RootElement == null)
            return new Result("Output document doesn't have a body.", -1);

        var body = document.MainDocumentPart.RootElement.FirstChild!;

        var defaultFontSize = document
            .MainDocumentPart
            .StyleDefinitionsPart!
            .Styles!
            .Descendants<DocDefaults>()
            .First()
            .RunPropertiesDefault!
            .RunPropertiesBaseStyle!
            .FontSize!
            .Val!;
        
        var paragraphs = body
            .Descendants<Paragraph>()
            .Select(p => (Paragraph)p.CloneNode(true))
            .ToArray();
        var currentCharacterIndex = 0;
        var insertBreak = false;

        for (var i = 0; i < paragraphs.Length; i++)
        {
            var paragraph = paragraphs[i];
            if (currentCharacterIndex == message.Length)
                continue;
            
            var newRuns = paragraph.Descendants<Run>()
                .Select(r => (Run)r.CloneNode(true))
                .ToList();
            for (var j = 0; j < newRuns.Count; j++)
            {
                var run = newRuns[j];
                if (currentCharacterIndex == message.Length)
                    break;

                var currentFontSizeRaw = (StringValue?)run.RunProperties?.FontSize?.Val?.Clone();
                var currentFontSize = int.Parse((currentFontSizeRaw ?? defaultFontSize)!);
                var currentFonts = (RunFonts?)run.RunProperties?.RunFonts?.CloneNode(true);
                var currentStyle = (RunStyle?)run.RunProperties?.RunStyle?.CloneNode(true);
                var value = run.InnerText;

                var currentCharUpper = char.ToUpper(message[currentCharacterIndex]);
                var currentCharUpperIndex = value.IndexOf(currentCharUpper);
                var currentCharLower = char.ToLower(message[currentCharacterIndex]);
                var currentCharLowerIndex = value.IndexOf(currentCharLower);
                
                var currentChar = currentCharUpperIndex != -1 && currentCharUpperIndex < currentCharLowerIndex
                    ? currentCharUpper
                    : currentCharLower;

                var split = value.Split(currentChar, 2);

                if (split.Length == 2)
                {
                    var children = insertBreak
                        ? new OpenXmlElement[]
                        {
                            new Break(),
                            new Text
                            {
                                Text = split[0],
                                Space = SpaceProcessingModeValues.Preserve
                            }
                        }
                        : new OpenXmlElement[]
                        {
                            new Text
                            {
                                Text = split[0],
                                Space = SpaceProcessingModeValues.Preserve
                            }
                        };
                    insertBreak = false;
                    newRuns[j] = new Run(children)
                    {
                        RunProperties = new RunProperties
                        {
                            FontSize = new FontSize
                            {
                                Val = (StringValue?)currentFontSizeRaw?.Clone(),
                            },
                            RunFonts = (RunFonts?)currentFonts?.CloneNode(true),
                            RunStyle = (RunStyle?)currentStyle?.CloneNode(true),
                        }
                    };
                    var letter = new Text
                    {
                        Text = currentChar.ToString(),
                        Space = SpaceProcessingModeValues.Preserve
                    };
                    var runProps = new RunProperties
                    {
                        FontSize = new FontSize
                        {
                            Val = (currentFontSize + _fontSizeModifier).ToString()
                        },
                        RunFonts = (RunFonts?)currentFonts?.CloneNode(true),
                        RunStyle = (RunStyle?)currentStyle?.CloneNode(true)
                    };
                    newRuns.Insert(j + 1, new Run(letter)
                    {
                        RunProperties = runProps,
                    });
                    newRuns.Insert(j + 2, new Run(new Text
                        {
                            Text = split[1],
                            Space = SpaceProcessingModeValues.Preserve
                        }
                    )
                    {
                        RunProperties = new RunProperties
                        {
                            FontSize = new FontSize
                            {
                                Val = (StringValue?)currentFontSizeRaw?.Clone(),
                            },
                            RunFonts = (RunFonts?)currentFonts?.CloneNode(true),
                            RunStyle = (RunStyle?)currentStyle?.CloneNode(true)
                        }
                    });


                    currentCharacterIndex++;
                    j++;
                }
                else insertBreak = true;
            }
            paragraphs[i].RemoveAllChildren<Run>();
            paragraphs[i].Append(newRuns);
        }

        if (currentCharacterIndex != message.Length)
            return new Result("Can't hide this message in the given document.", -2);

        body.RemoveAllChildren<Paragraph>();
        body.Append(paragraphs);

        return 0;
    }
}