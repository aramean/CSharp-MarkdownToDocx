using System;
using System.IO;
using Utility;

namespace Utility
{
    class Program
    {
        static void Main(string[] args)
        {
            string markdownText = "# Header 1\n## Header 2\n### Header 3";
            string outputFilePath = "Headers.docx";

            try
            {
                // Convert Markdown headers to DOCX
                Utility.ConvertHeadersToDocx(markdownText, outputFilePath);
                Console.WriteLine($"Markdown converted to DOCX successfully. File saved at: {Path.GetFullPath(outputFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}