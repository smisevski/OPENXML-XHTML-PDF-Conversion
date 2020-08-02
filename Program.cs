/**
*  Forcefully applying CSS dimensions of an A4 format pagesize. Considering Column Layout, TextBoxes, Header/Footer, Images, Positions
*/

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Drawing.Imaging;
using System.Drawing.Text;
using iTextSharp;
using iTextSharp.tool.xml;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.html;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Vml;
using OpenXmlPowerTools;
using OpenXmlPowerTools.HtmlToWml;
using PdfSharp;
using PdfSharp.Pdf;
using MigraDoc;
using MigraDoc.DocumentObjectModel; 
namespace OPENXML
{
    class Program
    {
        static void Main(string[] args)
        { 

            string resourceFilepath = args[0];
            string outputFilepath = args[1];

            byte[] byteArray = File.ReadAllBytes(resourceFilepath);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);

                using (WordprocessingDocument wpd = WordprocessingDocument.Open(ms, true))
                {
                    Body body = wpd.MainDocumentPart.Document.Body;
                    int imageCounter = 0;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        AdditionalCss = "body { width: 210mm!important;height: 100%;max-width: 210mm; padding: 0; background-color: beige; padding: 1cm;}",
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png") imageFormat = ImageFormat.Png;
                            else if (extension == "gif") imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            if (imageFormat == null) return null;

                            string base64 = null;
                            try
                            {
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    imageInfo.Bitmap.Save(ms, imageFormat);
                                    var ba = ms.ToArray();
                                    base64 = System.Convert.ToBase64String(ba);
                                }
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            { return null; }

                            ImageFormat format = imageInfo.Bitmap.RawFormat;
                            ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders()
                                                      .First(c => c.FormatID == format.Guid);
                            string mimeType = codec.MimeType;

                            string imageSource =
                                   string.Format("data:{0};base64,{1}", mimeType, base64);

                            XElement img = new XElement(Xhtml.img,
                                  new XAttribute(NoNamespace.src, imageSource),
                                  imageInfo.ImgStyleAttribute,
                                  imageInfo.AltText != null ?
                                       new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }

                    };

                    
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wpd, settings);
                    html.Save(outputFilepath);

                    Console.WriteLine("Done converting DOCX to XHTML.....!");



                    List<HeaderPart> headerPts = wpd.MainDocumentPart.HeaderParts.ToList();

                    DocumentFormat.OpenXml.Wordprocessing.SplitPageBreakAndParagraphMark pgBr = wpd.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.SplitPageBreakAndParagraphMark>().FirstOrDefault();

                    Console.WriteLine(pgBr.InnerXml);


                };


            }
        }

    }
}

