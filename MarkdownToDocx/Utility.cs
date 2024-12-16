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
                    else if (line.StartsWith("#### "))
                    {
                        AddHeaderToDocx(body, line.Substring(5), "Heading4");  // Standard H4 style
                    }
                    else
                    {
                        // For other lines, apply a generic Heading style
                        AddHeaderToDocx(body, line, "Generic Heading");  // Generic Heading style (not linked to markdown level)
                    }
                }
            }
        }

        // Add standard heading styles (Heading 1, Heading 2, Heading 3, Heading 4) to the style definitions
        private static void AddStandardHeadingStyles(Styles styles)
        {
            styles.Append(
                CreateHeadingStyle("Heading1", "Heading 1", JustificationValues.Left, true, false, false, 18),  // 18 pt font size
                CreateHeadingStyle("Heading2", "Heading 2", JustificationValues.Left, false, true, false, 16), // 16 pt font size
                CreateHeadingStyle("Heading3", "Heading 3", JustificationValues.Left, false, false, true, 14), // 14 pt font size
                CreateHeadingStyle("Heading4", "Heading 4", JustificationValues.Left, false, false, false, 12), // 12 pt font size
                CreateHeadingStyle("Generic Heading", "Generic Heading", JustificationValues.Left, false, false, false, 12) // Generic Heading style
            );
        }

        // Helper method to create a heading style
        private static Style CreateHeadingStyle(string styleId, string styleName, JustificationValues justification, bool bold, bool italic, bool underline, int fontSize)
        {
            // Create RunProperties based on passed boolean values
            var runProperties = new RunProperties();
            
            if (bold)
                runProperties.Append(new Bold());
            
            if (italic)
                runProperties.Append(new Italic());
            
            if (underline)
                runProperties.Append(new Underline() { Val = UnderlineValues.Single });

            runProperties.Append(new FontSize() { Val = (fontSize * 2).ToString() });  // Multiply by 2 for OpenXML (half-points)

            return new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                StyleName = new StyleName() { Val = styleName },
                BasedOn = new BasedOn() { Val = "Normal" },
                NextParagraphStyle = new NextParagraphStyle() { Val = "Normal" },
                StyleParagraphProperties = new StyleParagraphProperties(new Justification() { Val = justification }),
                StyleRunProperties = new StyleRunProperties(runProperties)
            };
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

            // Define the style based on the style ID (Heading1, Heading2, Heading3, Heading4)
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