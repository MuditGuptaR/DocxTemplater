using NUnit.Framework;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocxTemplater.Test
{
    [TestFixture]
    public class MergeFieldTests
    {
        [Test]
        public void SimpleMergeField_IsReplaced()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                var body = new Body();
                var paragraph = new Paragraph();
                var simpleField = new SimpleField() { Instruction = "MERGEFIELD my_simple_field" };
                var run = new Run(new Text("Default Text"));
                simpleField.Append(run);
                paragraph.Append(simpleField);
                body.Append(paragraph);
                mainPart.Document = new Document(body);
                wpDocument.Save();
            }

            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("my_simple_field", "Simple Field Value");
            var result = docTemplate.Process();

            result.Position = 0;
            using var resultDoc = WordprocessingDocument.Open(result, false);
            var resultText = resultDoc.MainDocumentPart.Document.Body.InnerText;
            Assert.That(resultText, Is.EqualTo("Simple Field Value"));
        }

        [Test]
        public void IfField_TrueCondition_ReplacesWithTrueText()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                var body = new Body();
                var paragraph = new Paragraph();
                var simpleField = new SimpleField() { Instruction = @"IF my_bool_field ""Is True"" ""Is False""" };
                var run = new Run(new Text("Default Text"));
                simpleField.Append(run);
                paragraph.Append(simpleField);
                body.Append(paragraph);
                mainPart.Document = new Document(body);
                wpDocument.Save();
            }

            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("my_bool_field", true);
            var result = docTemplate.Process();

            result.Position = 0;
            using var resultDoc = WordprocessingDocument.Open(result, false);
            var resultText = resultDoc.MainDocumentPart.Document.Body.InnerText;
            Assert.That(resultText, Is.EqualTo("Is True"));
        }

        [Test]
        public void IfField_FalseCondition_ReplacesWithFalseText()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                var body = new Body();
                var paragraph = new Paragraph();
                var simpleField = new SimpleField() { Instruction = @"IF my_bool_field ""Is True"" ""Is False""" };
                var run = new Run(new Text("Default Text"));
                simpleField.Append(run);
                paragraph.Append(simpleField);
                body.Append(paragraph);
                mainPart.Document = new Document(body);
                wpDocument.Save();
            }

            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("my_bool_field", false);
            var result = docTemplate.Process();

            result.Position = 0;
            using var resultDoc = WordprocessingDocument.Open(result, false);
            var resultText = resultDoc.MainDocumentPart.Document.Body.InnerText;
            Assert.That(resultText, Is.EqualTo("Is False"));
        }

        [Test]
        public void MixedContent_MergeFieldsAndPlaceholders()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                var body = new Body();
                var p1 = new Paragraph();
                var simpleField = new SimpleField() { Instruction = "MERGEFIELD my_simple_field" };
                p1.Append(simpleField);
                var p2 = new Paragraph(new Run(new Text(" and {{my_placeholder}}")));
                body.Append(p1);
                body.Append(p2);
                mainPart.Document = new Document(body);
                wpDocument.Save();
            }

            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("my_simple_field", "Hello");
            docTemplate.BindModel("my_placeholder", "World");
            var result = docTemplate.Process();

            result.Position = 0;
            using var resultDoc = WordprocessingDocument.Open(result, false);
            var resultText = resultDoc.MainDocumentPart.Document.Body.InnerText;
            Assert.That(resultText.Replace("\n", ""), Is.EqualTo("Hello and World"));
        }
    }
}
