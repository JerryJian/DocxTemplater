﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Templates.Images;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXml.Templates.Test
{
    internal class DocxTemplateTest
    {
        [Test]
        public void ReplaceTextBoldIsPreserved()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text("This Value:")),
                new Run(new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }), new Text("{{Property1}}")),
                new Run(new Text("will be replaced"))
            )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.AddModel("Property1", "Replaced");
            var result = docTemplate.Process();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open((Stream) result, false);
            var body = document.MainDocumentPart.Document.Body;
            // check that bold is preserved
            Assert.That(body.Descendants<Bold>().First().Val, Is.EqualTo(OnOffValue.FromBoolean(true)));
            // check that text is replaced
            Assert.That(body.Descendants<Text>().Skip(1).First().Text, Is.EqualTo("Replaced"));
        }

        [Test]
        public void BindCollection()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("{{PropertyInRoot}}")),
                    new Run(new Text("This Value:")),
                    new Run(new Text("{{#ds.Items}}"), new Text("For each run {{ds.Items.Name}}"))
                ),
            new Paragraph(
                    new Run(new Text("{{ds.Items.Value}}")),
                    new Run(new Text("{{#ds.Items.InnerCollection}}")),
                        new Run(new Text("{{ds.Items.Value}}")),
                        new Run(new Text("{{ds.Items.InnerCollection.Name}}")),
                        new Run(new Text("{{ds.Items.InnerCollection.InnerValue}}")),
                    new Run(new Text("{{/ds.Items.InnerCollection}}")),
                    new Run(new Text("{{/ds.Items}}")),
                    new Run(new Text("will be replaced"))
                )
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.AddModel("ds",
                new {
                        PropertyInRoot = "RootValue", 
                        Items = new[]
                        {
                            new {Name = " Item1 ", Value = " Value1 ", InnerCollection = new[]
                            {
                                new {Name = " InnerItem1 ", InnerValue = " InnerValue1 "}
                            }},
                            new {Name = " Item2 ", Value = " Value2 ", InnerCollection = new[]
                            {
                                new {Name = " InnerItem2 ", InnerValue = " InnerValue2 "},
                                new {Name = " InnerItem2a ", InnerValue = " InnerValue2b "}
                            }},
                        }
                });
            var result = docTemplate.Process();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            //check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("RootValueThis Value:For each run  Item1  Value1  Value1  InnerItem1  InnerValue1 For each run  Item2  Value2  Value2  InnerItem2  InnerValue2  Value2  InnerItem2a  InnerValue2b will be replaced"));
        }

        [Test]
        public void BindCollectionToTable()
        {
            var xml = @"<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">  
                      <w:tblPr>  
                        <w:tblW w:w=""5000"" w:type=""pct""/>  
                        <w:tblBorders>  
                          <w:top w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:left w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:bottom w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                          <w:right w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""auto""/>  
                        </w:tblBorders>  
                      </w:tblPr>  
                      <w:tblGrid>  
                        <w:gridCol w:w=""10296""/>  
                      </w:tblGrid>
                       <w:tr>  
                        <w:tc>  
                          <w:p><w:r><w:t>Header Col 1</w:t></w:r></w:p>  
                        </w:tc>
                        <w:tc>  
                          <w:p><w:r><w:t>Header Col 2</w:t></w:r></w:p>  
                        </w:tc>  
                      </w:tr>  
                      <w:tr>  
                        <w:tc>  
                          <w:tcPr>  
                            <w:tcW w:w=""0"" w:type=""auto""/>  
                          </w:tcPr>  
                          <w:p><w:r><w:t>{{#Items}}</w:t><w:t>{{Items.FirstVal}}</w:t></w:r></w:p>  
                        </w:tc>
                        <w:tc>  
                          <w:tcPr>  
                            <w:tcW w:w=""0"" w:type=""auto""/>  
                          </w:tcPr>  
                          <w:p><w:r><w:t>{{Items.SecondVal}}{{/Items}}</w:t></w:r></w:p>  
                        </w:tc>  
                      </w:tr>  
                    </w:tbl>";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Table(xml)));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.AddModel("ds",
                new
                {
                    Items = new[]
                    {
                        new {FirstVal = "CC_11", SecondVal = "CC_12"},
                        new {FirstVal = "CC_21", SecondVal = "CC_22"},
                    }
                });
            var result = docTemplate.Process();
            Assert.IsNotNull(result);
            result.Position = 0;
            //  result.SaveAsFileAndOpenInWord();
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var table = body.Descendants<Table>().First();
            var rows = table.Descendants<TableRow>().ToList();
            Assert.That(rows.Count, Is.EqualTo(3));
            Assert.That(rows[0].InnerText, Is.EqualTo("Header Col 1Header Col 2"));
            Assert.That(rows[1].InnerText, Is.EqualTo("CC_11CC_12"));
            Assert.That(rows[2].InnerText, Is.EqualTo("CC_21CC_22"));
        }

        [Test]
        public void ProcessBillTemplate()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            using var fileStream = File.OpenRead("Resources/BillTemplate.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.AddModel("ds", new
            {
                Company = new
                {
                    Logo = imageBytes
                },
                DisplayName = "John Doe",
                Street = "Main Street 42",
                City = "New York",
                Bills = new[]
                {
                    new
                    {
                        Date = DateTime.Now,
                        Name = "Rechnung für was",
                        CustomId = "R1045",
                        Amount = 1045.5m,
                        PaidAmount = 0m,
                        OpenAmount = 1045.5m
                    },
                    new
                    {
                        Date = DateTime.Now,
                        Name = "Bill 2",
                        CustomId = "R4242",
                        Amount = 1045.5m,
                        PaidAmount = 5042m,
                        OpenAmount = 1045.5m
                    },
                },
                Total = 1045.5m,
                TotalPaid = 0m,
                TotalOpen = 1045.5m,
                TotalDownPayment = 0m,
                HtmlTest = "<br class=\"k-br\"><table class=\"k-table\"><thead><tr style=\"height:19.85pt;\">" +
                           "<th colspan=\"2\" style=\"width:538px;border-width:1px;border-style:solid;border-color:#000000;background-color:#c1bfbf;vertical-align:middle;text-align:left;margin-left:60px;" +
                           "\">Document / Notes - This is table was generated from HTML</th></tr></thead><tbody><tr style=\"height:19.85pt;\">" +
                           "<td style=\"width: 538px;\" data-role=\"resizable\">Some Notes with special characters ä ö ü é and so on</td>" +
                           "<td style=\"width:162px;text-align:left;vertical-align:top;\">29.11.2023</td></tr></tbody></table><p>&#xFEFF;</p>"
            });

            var result = docTemplate.Process();
            result.Position = 0;
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Assert.That(paragraphs.Count, Is.EqualTo(63));
            // check some replacements
            Assert.That(body.InnerText.Contains("John Doe"), Is.EqualTo(true));
            Assert.That(body.InnerText.Contains("Main Street 42"), Is.EqualTo(true));
            Assert.That(body.InnerText.Contains("New York"), Is.EqualTo(true));

            // check table
            var table = body.Descendants<Table>().First();
            Assert.That(table.InnerText.Contains("Rechnung für was"), Is.EqualTo(true));
            Assert.That(table.InnerText.Contains("R1045"), Is.EqualTo(true));
            Assert.That(table.InnerText.Contains("Bill 2"), Is.EqualTo(true));
        }
    }


}
