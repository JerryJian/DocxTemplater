﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocxTemplater.Images;
#if NET6_0_OR_GREATER
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
#else
using System.Drawing;
using System.Drawing.Imaging;
#endif

namespace DocxTemplater.Test
{
    internal class ImageFormatterTest
    {
        [TestCase("jpg")]
        [TestCase("tiff")]
        [TestCase("png")]
        [TestCase("bmp")]
        [TestCase("gif")]
        public void ProcessTemplateWithDifferentImageTypes(string extension)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
#if NET6_0_OR_GREATER
            var img = Image.Load(imageBytes);
            Assert.That(img.Configuration.ImageFormatsManager.TryFindFormatByFileExtension(extension, out var format));
            var memStream = new MemoryStream();
            img.Save(memStream, format);
            imageBytes = memStream.ToArray();
#else
            var img = Image.FromStream(new MemoryStream(imageBytes));
            var memStream = new MemoryStream();
            img.Save(memStream, extension switch
            {
                "jpg" => ImageFormat.Jpeg,
                "tiff" => ImageFormat.Tiff,
                "png" => ImageFormat.Png,
                "bmp" => ImageFormat.Bmp,
                "gif" => ImageFormat.Gif,
                _ => ImageFormat.Jpeg,
            });
            imageBytes = memStream.ToArray();
#endif

            using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", new { MyLogo = imageBytes, EmptyArray = Array.Empty<byte>(), NullValue = (byte[])null });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        [TestCase("w:14cm,h:3cm")]
        [TestCase("w:14cm")]
        [TestCase("h:1cm, r:90")]
        [TestCase("w:1cm")]
        [TestCase("h:1cm")]
        public void InsertHugeImageInsertWithoutContainerFitsToPage(string argument)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            // change the size to be bigger than the page
#if NET6_0_OR_GREATER
            var img = Image.Load(imageBytes);
            img.Mutate(x => x.Resize(img.Width * 10, img.Height * 10));

            var bigImgStream = new MemoryStream();
            img.SaveAsJpeg(bigImgStream);
            imageBytes = bigImgStream.ToArray();
#else
            var img = Image.FromStream(new MemoryStream(imageBytes));
            var large = img.GetThumbnailImage(img.Width * 10, img.Height * 10, () => true, IntPtr.Zero);

            var bigImgStream = new MemoryStream();
            img.Save(bigImgStream, ImageFormat.Jpeg);
            imageBytes = bigImgStream.ToArray();
#endif

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(" + argument + ")}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }
    }
}
