using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xunit;

namespace DocumentFormat
{
    public class ReplaceTextSpec
    {
        [Fact]
        public void ShouldReplaceTextInWord()
        {
            var year = 2015;
            var month = 10;
            var day = 25;

            var document = @"Input\ReplaceText.docx";

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                docText = docText
                    .Replace("{year}", $"{year}")
                    .Replace("{month}", $"{month}")
                    .Replace("{day}", $"{day}");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}
