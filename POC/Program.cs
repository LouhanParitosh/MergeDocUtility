using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            string firstFile = "C:\\Personal\\Thinkpad to dell\\NAGP\\Forecia POC\\doc1.docx";
            string secondFile = "C:\\Personal\\Thinkpad to dell\\NAGP\\Forecia POC\\doc2.docx";
            string outputFile = "C:\\Personal\\Thinkpad to dell\\NAGP\\Forecia POC\\doc3.docx";

            MergeWordFiles(firstFile, secondFile, outputFile);

            Console.WriteLine("Files merged successfully.");
        }

        public static List<OpenXmlElement> ReadWordFileElements(string filePath)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                Body body = wordDoc.MainDocumentPart.Document.Body;
                foreach (var element in body.Elements())
                {
                    elements.Add(element.CloneNode(true)); // Cloning to avoid modifying the original document
                }
            }
            return elements;
        }

        public static void WriteWordFile(string filePath, List<OpenXmlElement> content, List<Style> styles)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                foreach (var element in content)
                {
                    body.AppendChild(element);
                }

                // Add styles
                if (styles.Count > 0)
                {
                    StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylePart.Styles = new Styles();
                    foreach (var style in styles)
                    {
                        stylePart.Styles.AppendChild(style.CloneNode(true));
                    }
                    stylePart.Styles.Save();
                }

                //Enable track changes
                DocumentSettingsPart settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
                settingsPart.Settings.AppendChild(new TrackRevisions());

                mainPart.Document.Save();
            }
        }

        public static void MergeWordFiles(string firstFilePath, string secondFilePath, string outputFilePath)
        {
            List<OpenXmlElement> firstFileContent = ReadWordFileElements(firstFilePath);
            List<OpenXmlElement> secondFileContent = ReadWordFileElements(secondFilePath);

            HashSet<string> uniqueContentHashes = new HashSet<string>();
            List<OpenXmlElement> mergedContent = new List<OpenXmlElement>();

            foreach (var element in firstFileContent)
            {
                string elementHash = element.OuterXml;
                if (uniqueContentHashes.Add(elementHash))
                {
                    mergedContent.Add(element.CloneNode(true));
                }
            }

            string mergedContentXml = ConvertToXmlString(mergedContent);

            // Set a breakpoint here and use the debugger to view mergedContentXml
            Console.WriteLine(mergedContentXml);
            Console.ReadKey();

            foreach (var element in secondFileContent)
            {
                string elementHash = element.OuterXml;
                if (uniqueContentHashes.Add(elementHash))
                {
                    mergedContent.Add(element.CloneNode(true));
                }
            }

            List<Style> styles = new List<Style>();
            using (WordprocessingDocument firstDoc = WordprocessingDocument.Open(firstFilePath, false))
            {
                styles.AddRange(GetStylesFromDocument(firstDoc));
            }

            using (WordprocessingDocument secondDoc = WordprocessingDocument.Open(secondFilePath, false))
            {
                styles.AddRange(GetStylesFromDocument(secondDoc));
            }

            WriteWordFile(outputFilePath, mergedContent, styles);
        }

        public static List<Style> GetStylesFromDocument(WordprocessingDocument doc)
        {
            List<Style> styles = new List<Style>();
            var stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            if (stylePart != null)
            {
                foreach (var style in stylePart.Styles.Elements<Style>())
                {
                    styles.Add((Style)style.CloneNode(true));
                }
            }
            return styles;
        }

        public static string ConvertToXmlString(List<OpenXmlElement> elements)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var element in elements)
            {
                sb.Append(element.OuterXml);
            }
            return sb.ToString();
        }
    }
}
