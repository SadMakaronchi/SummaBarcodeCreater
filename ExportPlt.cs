using Corel.Interop.VGCore;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using File = System.IO.File;
using Path = System.IO.Path;


namespace SummaBarcodeCreater
{
    public class ExportPlt
    {
        public Define_per define_per;
        public Settings settings;
        public LayerAndSearchContour layers;
        public Numbers numbers;
        public OPOS_marks opos;
        public ShapeRange obj;
        public ShapeRange obj2;
        public string plt1;
        public string plt2;
        public void define_obj(LayerAndSearchContour layers, Settings settings, Define_per define_per,Numbers numbers,OPOS_marks opos)
        {
            this.layers = layers;
            this.define_per = define_per;
            this.settings = settings;
            this.numbers = numbers;
            this.opos = opos;
            obj = layers.rzk.Shapes.All();
            obj.AddRange(layers.metk_sum.Shapes.All());
            define_per.ActPage.SizeHeight = 297; // нормирование страницы, если так не делать иногда корел смещает контур
            define_per.ActPage.SizeWidth = 210;
            double h = 297 / 25.4;
            double w = 210 / 25.4;
            var x = obj.CenterX;
            var y = obj.CenterY;
            obj.Rotate(270);
            obj.SetPositionEx(cdrReferencePoint.cdrBottomLeft, 0, 0);
            exp(define_per, settings,numbers.down_number.ToString(),h,w);
            obj.Rotate(-270);
            plt1 = AddParameterCut(obj, settings, opos, numbers.down_number.ToString());
            if (settings.barc2 == true)
            {
                obj2 = layers.rzk.Shapes.All();
                obj2.AddRange(layers.metk_sum.Shapes.All());
                obj2.Rotate(90);
                obj2.SetPositionEx(cdrReferencePoint.cdrBottomLeft, 0, 0);
                exp(define_per, settings, numbers.up_number.ToString(),h,w);
                obj2.Rotate(-90);
                plt2 = AddParameterCut(obj, settings, opos, numbers.up_number.ToString());
                obj2.CenterX = x;
                obj2.CenterY = y;
            }
            else
            {
                obj.CenterX = x;
                obj.CenterY = y;
            }
        }
        public void exp(Define_per define_per,Settings settings,string name,double h,double w)
        {
            string filename =  name + ".plt";
            string path_plt = Path.Combine(settings.path_plt, filename);
            layers.rzk.Shapes.All().CreateSelection();
            dynamic explt = define_per.corelDoc.ExportEx(path_plt, cdrFilter.cdrHPGL, cdrExportRange.cdrSelection);
            explt.PenLibIndex = 0;
            explt.FitToPage = false;
            explt.ScaleFactor = 100;
            explt.PageWidth = w;
            explt.PageHeight = h;
            explt.FillType = 0;
            explt.FillSpacing = 0.005;
            explt.FillAngle = 0;
            explt.HatchAngle = 90;
            explt.CurveResolution = 0.0001;
            explt.RemoveHiddenLines = true;
            explt.AutomaticWeld = false;
            explt.ExcludeWVC = true;
            explt.PlotterUnits = 1016;
            explt.PlotterOrigin = 1;   // левый нижний
            explt.Finish();

            define_per.corelDoc.ClearSelection();
        }
        public string AddParameterCut(ShapeRange obj,Settings settings,OPOS_marks opos,string name)
        {
            int mX = int.MaxValue;
            int mY = int.MaxValue;
            string filename = name + ".plt";
            string path_plt = Path.Combine(settings.path_plt, filename);
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
                if (l.StartsWith("PU") || l.StartsWith("PD")) // нормирование координат и исключение лишних комманд
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

            string plt = string.Join(Environment.NewLine, fixedLines);
            Order order = new Order();
            plt = order.Process(plt);
            using (StreamWriter sw = new StreamWriter(path_plt))
            {
                sw.WriteLine("\u001B;@:");
                sw.WriteLine("SET MARKER_X_DIS=" + (int)Math.Round(opos.step_y / 0.025) + ".");
                sw.WriteLine("SET MARKER_Y_DIS=" + (int)Math.Round((obj.SizeWidth - 3) / 0.025) + ".");
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
                sw.WriteLine("SET SMOOTHING=" + (settings.smothing == true? "ON":"OFF") + ".");
                sw.WriteLine("LOAD_MARKERS.");
                sw.WriteLine("END.");
                sw.WriteLine("IN;");
                sw.WriteLine("PA;");
                sw.WriteLine(plt);
                sw.WriteLine("PG;");
            }
            return plt;
        }
    }
}
