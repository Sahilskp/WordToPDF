using System;
using Microsoft.Office.Interop.Word;

public class Class1
{
	public static void main(string[] args)
	{
        var appWord = new Application();
        if (appWord.Documents != null)
        {
            //yourDoc is your word document
            var wordDocument = appWord.Documents.Open("C:\\Education\\3RD EVS Question paper - 28\\Aaliya.docx");
            string pdfDocName = "C:\\Education\\3RD EVS Question paper - 28\\Aaliya.pdf";
            if (wordDocument != null)
            {
                wordDocument.ExportAsFixedFormat(pdfDocName,
                WdExportFormat.wdExportFormatPDF);
                wordDocument.Close();
            }
            appWord.Quit();
        }
    }
}
