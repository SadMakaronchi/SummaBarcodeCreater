using Corel.Interop.VGCore;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using static System.Net.WebRequestMethods;
using corel = Corel.Interop.VGCore;
using File = System.IO.File;
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
        public void Open_pdf(long nbr, int n_met, int x_dis, int y_dis)
        {
            string path_settings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
            Settings_cut settings = new Settings_cut();
            using (FileStream fs = File.OpenRead(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                settings = (Settings_cut)xsz.Deserialize(fs);
            }
            int vel = settings.velosity;
            int over = settings.overcut;
            string smoth;
            string path_plt = settings.path_plt;
            if (settings.smothing == true)
            {
                smoth = "ON";
            }
            else
            {
                smoth = "OFF";
            }
            corelApp.ActiveDocument.Unit = cdrUnit.cdrInch;
            ShapeRange obj = corelApp.ActiveDocument.ActivePage.Shapes.All();
            obj.Rotate(270);
            double minX = obj.LeftX;
            double minY = obj.BottomY;
            obj.Move(-minX, -minY);
            string filename = nbr.ToString() + ".plt";
            path_plt = Path.Combine(path_plt, filename);
            dynamic explt = corelApp.ActiveDocument.ExportEx(path_plt, cdrFilter.cdrHPGL, cdrExportRange.cdrCurrentPage);
            explt.PenLibIndex = 0;
            explt.FitToPage = false;
            explt.ScaleFactor = 100;
            explt.PageWidth = corelApp.ActivePage.SizeWidth ;
            explt.PageHeight = corelApp.ActivePage.SizeHeight ;
            explt.FillType = 0;
            explt.FillSpacing = 0.005;
            explt.FillAngle = 0;
            explt.HatchAngle = 90;
            explt.CurveResolution = 0.0001;
            explt.RemoveHiddenLines = true;
            explt.AutomaticWeld = false;
            explt.ExcludeWVC = false;
            explt.PlotterUnits = 1016;
            explt.PlotterOrigin = 1;   // левый нижний
            explt.Finish();

            int mX = int.MaxValue;
            int mY = int.MaxValue;
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

            foreach (var l in lines)
            {
                if (l.StartsWith("PU") || l.StartsWith("PD"))
                {
                    var nums = Regex.Matches(l, @"-?\d+");
                    if (nums.Count >= 2)
                    {
                        int x = int.Parse(nums[0].Value);
                        int y = int.Parse(nums[1].Value);
                        mX = Math.Min(mX, x);
                        mY = Math.Min(mY, y);
                    }
                }
            }
            if (mX > 0) mX = 0;
            if (mY > 0) mY = 0;

            var fixedLines = lines.Select(l =>
            {
                if (l.StartsWith("PU") || l.StartsWith("PD"))
                {
                    var nums = Regex.Matches(l, @"-?\d+");
                    if (nums.Count >= 2)
                    {
                        int x = int.Parse(nums[0].Value) - mX;
                        int y = int.Parse(nums[1].Value) - mY;

                        return (l.StartsWith("PU") ? "PU" : "PD") + $"{x} {y};";
                    }
                }
                return l;
            })
                .ToArray();



           

            string plt = string.Join(Environment.NewLine,fixedLines);
            
            

            using (StreamWriter sw = new StreamWriter(path_plt))
            {
                sw.WriteLine("\u001B;@:");
                sw.WriteLine("SET MARKER_X_DIS="+x_dis+".");
                sw.WriteLine("SET MARKER_Y_DIS="+y_dis+".");
                sw.WriteLine("SET MARKER_X_SIZE=120.");
                sw.WriteLine("SET MARKER_Y_SIZE=120.");
                sw.WriteLine("SET MARKER_X_N="+n_met+".");
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

            corelApp.ActiveDocument.Unit = cdrUnit.cdrMillimeter;


        }
    }
}

        
    

