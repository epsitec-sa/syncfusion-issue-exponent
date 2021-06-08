﻿using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Exponent
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileToConvert = args[0];

            Program.Convert (fileToConvert);

            System.Console.WriteLine ("Conversion DONE");
        }

        private static void Convert(string docxPath)
        {
            using var docStream = new FileStream(docxPath, FileMode.Open);
            using var wordDocument = new WordDocument();
            wordDocument.Open (docStream, Syncfusion.DocIO.FormatType.Docx);

            using var render = new DocIORenderer();
            var pdfDocument = render.ConvertToPDF(wordDocument);
            
            var pdfPath = System.IO.Path.GetDirectoryName(docxPath);
            var pdfName = System.IO.Path.GetFileNameWithoutExtension(docxPath);
            var pdfFileName = System.IO.Path.Combine(pdfPath, $"{pdfName}.pdf");

            System.IO.File.Delete (pdfFileName);

            using var outputStream = new FileStream(pdfFileName, FileMode.CreateNew, FileAccess.Write);
            pdfDocument.Save (outputStream);
            pdfDocument.Close ();
        }
    }
}