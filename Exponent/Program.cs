using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Exponent
{
    class Program
    {
        static void Main(string[] args)
        {
            var folderToConvert = args.Length == 0 ? "..\\..\\..\\.." : args[0];

            var folderPath = System.IO.Path.GetFullPath (folderToConvert);
            var files = System.IO.Directory.GetFiles (folderPath, "*.docx");

            foreach (var file in files)
            {
                Program.Convert (file);
                
                System.Console.WriteLine ($"File '{file}' converted to PDF");
            }
        }

        private static void Convert(string docxPath)
        {
            using var docStream = new System.IO.FileStream (docxPath, System.IO.FileMode.Open);
            using var wordDocument = new WordDocument ();
            wordDocument.Open (docStream, Syncfusion.DocIO.FormatType.Docx);

            using var render = new DocIORenderer ();
            var pdfDocument = render.ConvertToPDF (wordDocument);

            var pdfPath = System.IO.Path.GetDirectoryName (docxPath);
            var pdfName = System.IO.Path.GetFileNameWithoutExtension (docxPath);
            var pdfFileName = System.IO.Path.Combine (pdfPath, $"{pdfName}.pdf");

            System.IO.File.Delete (pdfFileName);

            using var outputStream = new System.IO.FileStream (pdfFileName, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write);
            pdfDocument.Save (outputStream);
            pdfDocument.Close ();
        }
    }
}
