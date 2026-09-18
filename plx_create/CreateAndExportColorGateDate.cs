using Corel.Interop.VGCore;
using SummaBarcodeCreater;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Xml.Linq;
using Bitmap = System.Drawing.Bitmap;
using Path = System.IO.Path;


public class CreateAndExportColorGateDate
{
    public string base64;
    public string date;
    public Bitmap bitmap;
    public Define_per define_per;
    public LayerAndSearchContour layers;
    public Settings settings;
    public ExportPlt export_plt;
    public int x;
    public int y;
    public bool rotate;
    public string hpgl_string;
    public string Plx(Define_per define_per,LayerAndSearchContour layers,Settings settings,bool rotate,string hpgl)
    {
        this.define_per = define_per;
        this.layers = layers;
        this.settings = settings;
        this.rotate = rotate;
        Assembly assembly = typeof(CreateAndExportColorGateDate).Assembly;
        Stream stream = assembly.GetManifestResourceStream("SummaBarcodeCreater.plx_create.ColorGateData.xml");
        layers.prnt.Shapes.All().CreateSelection();
        layers.metk_sum.Shapes.All().AddToSelection();
        base64 = BitmapToBase64();
        define_per.corelDoc.ClearSelection();
        date = DateTimeOffset.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
        var curve = HpglToXml.Convert(hpgl);
        MessageBox.Show(curve.ToString());
        XDocument doc = XDocument.Load(stream);

        XNamespace ns = "urn:schemas-colorgate-com:cutjobinformation";

        // JobInfo
        XElement jobInfo = doc.Root.Element(ns + "JobInfo");
        // Thumbnail
        XElement thumbnail = jobInfo.Element(ns + "Thumbnail");
                 thumbnail.SetAttributeValue("Data", base64);
        //JobName
        XElement JobName = jobInfo.Element(ns + "JobName");
        JobName.Value = define_per.DocName;
        // Printed
        XElement printed = jobInfo.Element(ns + "Printed");
                 printed.Value = date;
        // Created
        XElement created = jobInfo.Element(ns + "Created");
                 created.Value = date;
        // PreviewData -> PreviewPage -> Images -> Image
        XElement image = doc.Root.Element(ns + "PreviewData")
                                 .Element(ns + "PreviewPage")
                                 .Element(ns + "Images")
                                 .Element(ns + "Image");
                 image.SetAttributeValue("Data", base64);

        XElement curvedat = doc.Root.Element(ns + "PreviewData")
                                 .Element(ns + "PreviewPage")
                                 .Element(ns + "Paths")
                                 .Element(ns + "Path");
        curvedat.RemoveNodes();

        curvedat.Add(curve.Nodes());

       

        XElement Size = doc.Root.Element(ns + "PreviewData")
                                .Element(ns + "PreviewPage")
                                .Element(ns + "Images")
                                .Element(ns + "Image")
                                .Element(ns + "Position");
                 Size.SetAttributeValue("Width", layers.metk_sum.Shapes.All().SizeWidth);
                 Size.SetAttributeValue("Height", layers.metk_sum.Shapes.All().SizeHeight);
        


        // Возвращаем XML в string
        return doc.ToString();
    }

    public string BitmapToBase64()
    {
        string tempFile = Path.Combine(settings.path_plt,"image.jpg");
        try
        {
            double minStr = Math.Min(define_per.corelDoc.Selection().SizeWidth, define_per.corelDoc.Selection().SizeHeight);
            if(minStr == define_per.corelDoc.Selection().SizeWidth)
            {
                y = 3000;
                x = (int)(minStr / define_per.corelDoc.Selection().SizeHeight * 3000);
            }
            else
            {
                x = 3000;
                y = (int)(minStr / define_per.corelDoc.Selection().SizeWidth * 3000);
            }
            dynamic exportFilter = define_per.corelDoc.ExportBitmap(tempFile, cdrFilter.cdrJPEG, cdrExportRange.cdrSelection, cdrImageType.cdrRGBColorImage, x, y, 96, 96);
            exportFilter.Finish();
            if (rotate == true)
            {
                Bitmap bitmap = new Bitmap(tempFile);
                bitmap.RotateFlip(RotateFlipType.Rotate180FlipNone);
                bitmap.Save(tempFile, ImageFormat.Jpeg);
            }
            byte[] bytes = File.ReadAllBytes(tempFile);
            return Convert.ToBase64String(bytes);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }
}