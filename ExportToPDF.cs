using Corel.Interop.VGCore;
using iText.Kernel.Pdf;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using static System.Net.WebRequestMethods;
using Application = Corel.Interop.VGCore.Application;
namespace SummaBarcodeCreater
{
    public class ExportToPDF
    {
        private const double HpglUnitsPerMm = 40.0;
        private const double PointsPerInch = 72.0;
        private const double MmPerInch = 25.4;
        public Application corelApp;
        public Numbers num;
        public Settings set;
        public Define_per perm;
        public LayerAndSearchContour lrs;
        public void BeginExport(Application app, Numbers numbers,
                                                 Settings settings,
                                                 Define_per define_per,
                                                 LayerAndSearchContour layers)
        {
            corelApp = app;
            num = numbers;
            perm = define_per;
            lrs = layers;
            string path = settings.path_plt;
            string name = Path.Combine(path, numbers.down_number.ToString() + ".pdf");
            lrs.metk_sum.Shapes.All().CreateSelection();
            lrs.rzk.Shapes.All().AddToSelection();
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


            var contours = PdfContourExtractor.Extract(name,"CutContour"); //получение контуров из pdf
                                                                           // Получаем ширину именно той страницы, которую извлекали.
            double pageWidthPoints;

            using (var reader = new iText.Kernel.Pdf.PdfReader(name))
            using (var pdf = new iText.Kernel.Pdf.PdfDocument(reader))
            {
                pageWidthPoints = pdf.GetPage(1).GetPageSize().GetWidth() - 0.7 * 72.0 / 25.4; //ручное вычитание в пунктах из-за приколов корела
            }

            PltWriter.Write(Path.Combine(path, numbers.down_number.ToString() + ".plt"), contours,pageWidthPoints);

        }



    }
}
