using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Utility
{
    public class Utility
    {
        public static void ConvertHeadersToDocx(string markdownText, string outputFilePath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputFilePath, WordprocessingDocumentType.Document))
            {
                // Create the main document part
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Add style definitions (styles will be standardized)
                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = new Styles();

                // Ensure the standard heading styles are available
                AddStandardHeadingStyles(stylePart.Styles);

                // Convert Markdown to DOCX and add headers
                var body = mainPart.Document.Body;

                // Parse markdown for headers (simplified example)
                var headers = markdownText.Split('\n');

                foreach (var line in headers)
                {
                    if (line.StartsWith("# "))
                    {
                        AddHeaderToDocx(body, line.Substring(2), "Heading1");  // Standard H1 style
                    }
                    else if (line.StartsWith("## "))
                    {
                        AddHeaderToDocx(body, line.Substring(3), "Heading2");  // Standard H2 style
                    }
                    else if (line.StartsWith("### "))
                    {
                        AddHeaderToDocx(body, line.Substring(4), "Heading3");  // Standard H3 style
                    }
                }
            }
        }

        // Add standard heading styles (Heading 1, Heading 2, Heading 3) to the style definitions
        private static void AddStandardHeadingStyles(Styles styles)
        {
            styles.Append(
                new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1",  // Standard built-in style for Heading 1
                    StyleName = new StyleName() { Val = "Heading 1" },  // Localizable name, but it's safe in this context
                    BasedOn = new BasedOn() { Val = "Normal" },
                    NextParagraphStyle = new NextParagraphStyle() { Val = "Normal" },
                    StyleParagraphProperties = new StyleParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    ),
                    StyleRunProperties = new StyleRunProperties(new Bold())
                },
                new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading2",  // Standard built-in style for Heading 2
                    StyleName = new StyleName() { Val = "Heading 2" },  // Localizable name, but it's safe in this context
                    BasedOn = new BasedOn() { Val = "Normal" },
                    NextParagraphStyle = new NextParagraphStyle() { Val = "Normal" },
                    StyleParagraphProperties = new StyleParagraphProperties(),
                    StyleRunProperties = new StyleRunProperties(new Italic())
                },
                new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading3",  // Standard built-in style for Heading 3
                    StyleName = new StyleName() { Val = "Heading 3" },  // Localizable name, but it's safe in this context
                    BasedOn = new BasedOn() { Val = "Normal" },
                    NextParagraphStyle = new NextParagraphStyle() { Val = "Normal" },
                    StyleParagraphProperties = new StyleParagraphProperties(),
                    StyleRunProperties = new StyleRunProperties(new Underline() { Val = UnderlineValues.Single })
                }
            );
        }

        // Method to add headers to the document
        public static void AddHeaderToDocx(Body body, string text, string styleId)
        {
            // Ensure 'body' is not null
            if (body == null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            // Create a paragraph to hold the header
            Paragraph paragraph = new Paragraph();

            // Define the style based on the style ID (Heading1, Heading2, Heading3)
            ParagraphProperties paraProps = new ParagraphProperties();
            paraProps.Append(new ParagraphStyleId() { Val = styleId });  // Apply the style using the ID

            // Apply paragraph properties
            paragraph.Append(paraProps);

            // Create a run to hold the actual text
            Run run = new Run(new Text(text));

            // Append the run to the paragraph
            paragraph.Append(run);

            // Finally, add the paragraph to the body
            body.Append(paragraph);
        }
    }
}