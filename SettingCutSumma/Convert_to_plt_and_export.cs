using Corel.Interop.VGCore;
using SummaMetki.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;
using System.Xml.Serialization;
using static System.Net.Mime.MediaTypeNames;
using corel = Corel.Interop.VGCore;
using Path = System.IO.Path;


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
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
            Settings_cut settings = new Settings_cut();
            using (FileStream fs = File.OpenRead(path))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                xsz.Serialize(fs, settings);
            }
            string vel = $"\"{settings.velosity}\"";
            string over = $"\"{settings.overcut}\"";
            string smoth;
            string path_plt = $"\"{settings.path_plt}\"";
            if (settings.smothing == true)
            {
                smoth = "ON";
            }
            else
            {
                smoth = "OFF";
            }

            ShapeRange obj = corelApp.ActiveDocument.ActivePage.Shapes.All();
            obj.Rotate(270);
            obj.CreateSelection();
            obj.SetPositionEx(cdrReferencePoint.cdrBottomLeft, 0, 0);
            dynamic explt = corelApp.ActiveDocument.ExportEx(path_plt, cdrFilter.cdrHPGL, cdrExportRange.cdrSelection);
            explt.PenLibIndex = 0;
            explt.FitToPage = false;
            explt.ScaleFactor = 100;
            explt.PageWidth = corelApp.ActivePage.SizeWidth / 25.4;
            explt.PageHeight = corelApp.ActivePage.SizeHeight / 25.4;
            explt.FillType = 0;
            explt.FillSpacing = 0.005;
            explt.FillAngle = 0;
            explt.HatchAngle = 90;
            explt.CurveResolution = 0.0001;
            explt.AutomaticWeld = false;
            explt.ExcludeWVC = true;
            explt.PlotterUnits = 1016;
            explt.PlotterOrigin = 1;   // левый нижний
            explt.Finish();

           
            var lines = File.ReadAllLines(path_plt)
                            .Where(l =>
                            {
                                string s = l.TrimStart();
                                return !(
                                    s.StartsWith("SP") ||
                                    s.StartsWith("IN") ||
                                    s.StartsWith("LT")
                                );
                            })

                            .ToArray();

            string plt = string.Join(Environment.NewLine, lines);
            
            

            using (StreamWriter sw = new StreamWriter(path_plt))
            {
                sw.WriteLine("\u001B;@:");
                sw.WriteLine("SET MARKER_X_DIS=.");
                sw.WriteLine("SET MARKER_Y_DIS=.");
                sw.WriteLine("SET MARKER_X_SIZE=120.");
                sw.WriteLine("SET MARKER_Y_SIZE=120.");
                sw.WriteLine("SET MARKER_X_N=.");
                sw.WriteLine("SET SPECIAL_LOAD=OPOS_XY.");
                sw.WriteLine("SET VELOCITY="+vel+".");
                sw.WriteLine("SET OVERCUT="+over+".");
                sw.WriteLine("SET SMOOTHING="+smoth+".");
                sw.WriteLine("LOAD_MARKERS.");
                sw.WriteLine("END.");
                sw.WriteLine("IN;");
                sw.WriteLine("PA;");
                sw.WriteLine(plt);
                sw.WriteLine("PG;");
            }


        }
    }
}

        
    

