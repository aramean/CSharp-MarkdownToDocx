using System;
using System.IO;
using Utility;

namespace Utility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure the input markdown file path is passed as an argument
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: Utility <markdown-file-path>");
                return;
            }

            string markdownFilePath = args[0];  // Input markdown file path
            string outputFilePath = "Headers.docx";  // Output DOCX file path

            try
            {
                // Read the markdown content from the input file
                if (!File.Exists(markdownFilePath))
                {
                    Console.WriteLine($"Error: The file '{markdownFilePath}' does not exist.");
                    return;
                }

                string markdownText = File.ReadAllText(markdownFilePath);

                // Convert Markdown headers to DOCX
                Utility.ConvertHeadersToDocx(markdownText, outputFilePath);

                // Output the result
                Console.WriteLine($"Markdown converted to DOCX successfully. File saved at: {Path.GetFullPath(outputFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}