using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxStenography.Core;

public class Stenographer
{
    private readonly int _fontSizeModifier;

    public Stenographer(bool debugMode = false)
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

                var currentFontSize = int.Parse((run.RunProperties?.FontSize?.Val ?? defaultFontSize)!);
                var value = run.InnerText;

                var currentCharUpper = char.ToUpper(message[currentCharacterIndex]);
                var currentCharUpperIndex = value.IndexOf(currentCharUpper);
                var currentCharLower = char.ToLower(message[currentCharacterIndex]);
                var currentCharLowerIndex = value.IndexOf(currentCharUpper);
                
                var currentChar = currentCharUpperIndex != -1 && currentCharUpperIndex < currentCharLowerIndex
                    ? currentCharUpper
                    : currentCharLower;

                var split = value.Split(currentChar, 2);

                if (split.Length == 2)
                {
                    newRuns[j] = new Run(new Text(split[0]));
                    var letter = new Text(currentChar.ToString());
                    var runProps = new RunProperties
                    {
                        FontSize = new FontSize
                        {
                            Val = (currentFontSize + _fontSizeModifier).ToString()
                        }
                    };
                    newRuns.Insert(j + 1, new Run(letter)
                    {
                        RunProperties = runProps,
                    });
                    newRuns.Insert(j + 2, new Run(new Text(split[1])));
                    currentCharacterIndex++;
                    j++;
                }
            }

            paragraphs[i] = new Paragraph(newRuns);
        }

        if (currentCharacterIndex != message.Length)
            return new Result("Can't hide this message in the given document.", -2);

        body.RemoveAllChildren<Paragraph>();
        body.Append(paragraphs);

        return 0;
    }
}