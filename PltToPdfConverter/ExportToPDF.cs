using Corel.Interop.VGCore;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using System;
using System.IO;
using Application = Corel.Interop.VGCore.Application;
using Path = System.IO.Path;
using Rectangle = iText.Kernel.Geom.Rectangle;

namespace SummaBarcodeCreater
{
    public class ExportToPDF
    {
        public Application corelApp;
        public Numbers num;
        public Settings set;
        public Define_per perm;
        public LayerAndSearchContour lrs;
        public string PltOrdered;
        public double pageWidthPoints;
        public double pageHeightPoints;
        public string plt;
        public string BeginExport(Application app, long numbName,
                                                 Settings settings,
                                                 Define_per define_per,
                                                 LayerAndSearchContour layers,
                                                 OPOS_marks opos,
                                                 bool rotate)
        {
            corelApp = app;
            perm = define_per;
            lrs = layers;
            string path = settings.path_plt;
            string name = Path.Combine(path, numbName.ToString() + ".pdf");
            lrs.metk_sum.Shapes.All().CreateSelection();
            lrs.rzk.Shapes.All().AddToSelection();

            if (settings.pdf_contour && rotate)
                opos.mark_bottom.Selected = false;
            if (settings.pdf_contour && !rotate)
                opos.mark_top.Selected = false;

            var exprt = perm.corelDoc.PDFSettings;
            exprt.PublishRange = pdfExportRange.pdfSelection;
            exprt.PageRange = "1";
            exprt.Author = "SummaBarcodeCreater";
            exprt.Subject = "";
            exprt.Keywords = "";
            exprt.BitmapCompression = pdfBitmapCompressionType.pdfLZW;
            exprt.JPEGQualityFactor = 10;
            exprt.TextAsCurves = true;
            exprt.EmbedFonts = false;
            exprt.EmbedBaseFonts = true;
            exprt.TrueTypeToType1 = true;
            exprt.SubsetFonts = true;
            exprt.SubsetPct = 80;
            exprt.CompressText = false;
            exprt.Encoding = pdfEncodingType.pdfBinary;
            exprt.DownsampleColor = false;
            exprt.DownsampleGray = false;
            exprt.DownsampleMono = false;
            exprt.ColorResolution = 200;
            exprt.MonoResolution = 600;
            exprt.GrayResolution = 200;
            exprt.Hyperlinks = false;
            exprt.Bookmarks = false;
            exprt.Thumbnails = false;
            exprt.Startup = 0;
            exprt.ComplexFillsAsBitmaps = false;
            exprt.Overprints = true;
            exprt.Halftones = false;
            exprt.MaintainOPILinks = false;
            exprt.FountainSteps = 256;
            exprt.EPSAs = 0;
            exprt.pdfVersion = 0;
            exprt.IncludeBleed = false;
            exprt.Bleed = 31750;
            exprt.Linearize = false;
            exprt.CropMarks = false;
            exprt.RegistrationMarks = false;
            exprt.DensitometerScales = false;
            exprt.FileInformation = false;
            exprt.ColorMode = pdfColorMode.pdfNative;
            exprt.ColorProfile = pdfColorProfile.pdfSeparationProfile;
            exprt.EmbedFilename = "";
            exprt.EmbedFile = false;
            exprt.JP2QualityFactor = 10;
            exprt.TextExportMode = 0;
            exprt.PrintPermissions = 0;
            exprt.EditPermissions = 0;
            exprt.ContentCopyingAllowed = false;
            exprt.OpenPassword = "";
            exprt.PermissionPassword = "";
            exprt.EncryptType = pdfEncryptionType.pdfEncryptTypeAES;
            exprt.OutputSpotColorsAs = 0;
            exprt.OverprintBlackLimit = 95;
            exprt.ProtectedTextAsCurves = true;
            exprt.UsePageBoundingBox = false;
            perm.corelDoc.PublishToPDF(name);

            if (!settings.pdf_contour)
            {
                foreach (string color in settings.color_name)
                {
                    var contours = PdfContourExtractor.Extract(name, color); //получение контуров из pdf
                    using (var reader = new PdfReader(name))
                    using (var pdf = new PdfDocument(reader))
                    {
                        pageWidthPoints = pdf.GetPage(1).GetPageSize().GetWidth() - 0.7 * 72.0 / 25.4; //ручное вычитание в пунктах из-за приколов корела
                        pageHeightPoints = pdf.GetPage(1).GetPageSize().GetHeight() - 0.7 * 72.0 / 25.4;
                    }
                    plt = PltWriter.Write(contours, pageWidthPoints, pageHeightPoints, rotate);
                    Order order = new Order();
                    PltOrdered = order.Process(plt.ToString());
                }

                using (StreamWriter sw = new StreamWriter(Path.Combine(path, numbName.ToString() + ".plt")))
                {
                    sw.WriteLine("\u001B;@:");
                    sw.WriteLine("SET MARKER_X_DIS=" + (int)Math.Round(opos.step_y / 0.025) + ".");
                    sw.WriteLine("SET MARKER_Y_DIS=" + ((int)Math.Round(pageWidthPoints * 25.4 / 72.0 * 40.0) - 120) + ".");
                    sw.WriteLine("SET MARKER_X_SIZE=120.");
                    sw.WriteLine("SET MARKER_Y_SIZE=120.");
                    sw.WriteLine("SET MARKER_X_N=" + (opos.metk_y + 1) + ".");
                    sw.WriteLine("SET SPECIAL_LOAD=OPOS_XY.");
                    if (settings.velosity != -1)
                    {
                        sw.WriteLine("SET VELOCITY=" + settings.velosity + ".");
                    }
                    if (settings.overcut != -1)
                    {
                        sw.WriteLine("SET OVERCUT=" + settings.overcut + ".");
                    }
                    sw.WriteLine("SET SMOOTHING=" + (settings.smothing == true ? "ON" : "OFF") + ".");
                    sw.WriteLine("LOAD_MARKERS.");
                    sw.WriteLine("END.");
                    sw.WriteLine("IN;");
                    sw.WriteLine("PA;");
                    sw.WriteLine(PltOrdered);
                    sw.WriteLine("PG;");

                    File.Delete(name);
                }

                return PltOrdered;
            }
            else
            {
                if (rotate)
                {
                    string tempName = Path.Combine(path,numbName.ToString() + "_temp.pdf");
                    RotatePdfContent(name, tempName);
                    File.Delete(name);
                    File.Move(tempName, name);
                }

                return null;
            }


                
        }
        private void RotatePdfContent(string sourcePath, string outputPath)
        {
            using (var reader = new PdfReader(sourcePath))
            using (var sourcePdf = new PdfDocument(reader))
            using (var writer = new PdfWriter(outputPath))
            using (var destPdf = new PdfDocument(writer))
            {
                PdfPage sourcePage = sourcePdf.GetPage(1);

                Rectangle size = sourcePage.GetPageSize();
                float width = size.GetWidth();
                float height = size.GetHeight();

                // Создаём новую страницу исходного размера
                PdfPage destPage = destPdf.AddNewPage(
                    new iText.Kernel.Geom.PageSize(width, height)
                );

                // Копируем содержимое исходной страницы
                PdfFormXObject form =
                    sourcePage.CopyAsFormXObject(destPdf);

                // Физически поворачиваем содержимое на 180°
                PdfCanvas canvas = new PdfCanvas(destPage);

                canvas.AddXObjectWithTransformationMatrix(
                    form,
                    -1, 0,
                     0, -1,
                     width, height
                );

                // Поворот задаётся содержимым, не метаданными
                destPage.SetRotation(0);
            }
        }



    }
}
