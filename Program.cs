using System;

/* Open XML SDK */
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLSDK
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("test.docx", true))
            {
                /* Read Text in stream */
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                /* Replace Text with regex expression */
                Regex regexText = new Regex("ReplaceMe");
                docText = regexText.Replace(docText, "Replaced Successfully!");

                /* Write Text into stream */
                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream()))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}
