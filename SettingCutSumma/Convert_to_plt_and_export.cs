using Corel.Interop.VGCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Graphics.Operations.PathConstruction;
using static UglyToad.PdfPig.Geometry.PdfPath;
using corel = Corel.Interop.VGCore;

namespace SummaMetki
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("0dbf8ab4-7de3-4572-a4eb-e918c3ef77d2")]

    public class Convert_to_plt_and_export
    {
        public corel.Application corelApp;
        public void Init(corel.Application app)
        {
            corelApp = app;
        }
        public void Open_pdf()
        {
            ShapeRange obj = corelApp.ActiveDocument.ActivePage.Shapes.All();
            obj.Rotate(270);
            obj.CreateSelection();
            dynamic expflt = corelApp.ActiveDocument.ExportEx("output.plt", cdrFilter.cdrHPGL, cdrExportRange.cdrSelection);
            expflt.PenLibIndex = 0;
            expflt.FitToPage = false;
            expflt.ScaleFactor = 100.0;

            

            expflt.FillType = 0;
            expflt.FillSpacing = 0.005;
            expflt.FillAngle = 0.0;
            expflt.HatchAngle = 90.0;

            expflt.CurveResolution = 0.0001;
            expflt.RemoveHiddenLines = false;
            expflt.AutomaticWeld = false;
            expflt.ExcludeWVC = true;

            expflt.PlotterUnits = 1016; // HPGL (40 units/mm)
            expflt.PlotterOrigin = 1;   // левый нижний
            expflt.Finish();
        }
    }
}
        
    

