using Corel.Interop.VGCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using corel = Corel.Interop.VGCore;
using System.Xml.Serialization;
using System.IO;

namespace SummaMetki
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("b3f7d9a3-02f6-4259-9b13-28c2c8070cfa")]

    public class Settings_cut
    {
        public string path_plt { get; set; }
        public int velosity { get; set; } = 0;
        public bool barc2 { get; set; }
        public double overcut { get; set; }
        public bool smothing { get; set; }
    }
    public class Entry
    {
        public corel.Application corelApp;
        public void Init(corel.Application app)
        {
            corelApp = app;
        }
        


        public corel.Application crl = new corel.Application();
        public void Start()
        {
            
            
                
            

            
            var win = new MainWindow(crl);
            win.Show();

        }
        private Thread uiThread;
        private progress win;
        public void Progress()
        {
            if (uiThread != null && uiThread.IsAlive)
            {
                win.Dispatcher.Invoke(() =>
                {
                    if (win.WindowState == WindowState.Minimized)
                        win.WindowState = WindowState.Normal;
                    win.Activate();
                });
                return;
            }
            uiThread = new Thread(() =>
            {

                win = new progress();
                win.Closed += (s, e) => win.Dispatcher.InvokeShutdown();
                win.Show();
                System.Windows.Threading.Dispatcher.Run();
            });

            uiThread.SetApartmentState(ApartmentState.STA); // обязательно для WPF
            uiThread.IsBackground = true;
            uiThread.Start();
        }



        public void Progress_stop()
        {
            if (win != null)
            {
                win.Dispatcher.Invoke(() =>
                {
                    win.Close();
                    win = null;
                });
            }
        }

        public void Begin()
        {
            if (corelApp == null)
                throw new InvalidOperationException("Corel не инициализирован");
            if (corelApp.ActiveDocument == null)
                return;


            string name = corelApp.ActiveDocument.Name;
            corelApp.ActiveDocument.BeginCommandGroup("Добавление меток и баркодов для " + name);
            corelApp.ActiveDocument.Unit = corel.cdrUnit.cdrMillimeter;
            Boolean barc2 = true;

            //тут создаём рабочие слои

            Layer prnt = corelApp.ActiveDocument.ActivePage.Layers.Find("печать");
            if (prnt == null)
            {
                corelApp.ActiveLayer.Name = "печать";
                prnt = corelApp.ActiveDocument.ActivePage.Layers["печать"];
            }

            Layer rzk = corelApp.ActiveDocument.ActivePage.Layers.Find("резка");
            if (rzk == null)
            {
                corelApp.ActiveDocument.ActivePage.CreateLayer("резка");
                rzk = corelApp.ActiveDocument.ActivePage.Layers["резка"];
            }
            //ищем контур по имени спотового цвета абриса
            ShapeRange cut = corelApp.ActiveDocument.ActivePage.Shapes.FindShapes(Query: "@outline.color.name='CutContour'");
            cut.MoveToLayer(rzk);
            foreach (Shape s in cut)
            {
                s.BreakApart();
            }
            ShapeRange cut_next = rzk.Shapes.All();
            cut_next.Sort("@shape1.CenterY * 100 - @shape1.CenterX < @shape2.CenterY * 100 - @shape2.CenterX");
            Shape cutfnsh = cut_next.Combine();
            cutfnsh.BreakApart();

            prnt.Shapes.All().CreateSelection();
            rzk.Shapes.All().AddToSelection();
            ShapeRange r = corelApp.ActiveSelectionRange;

            int metk_y = (int)Math.Ceiling((r.SizeHeight + 6) / 400);
            int b2 = 0;
            if (barc2 == true)
            {
                b2 = 9;
            }
            double step_y = (r.SizeHeight + 20 + b2) / metk_y;
            if (step_y > 410)
            {
                metk_y = (int)Math.Ceiling((r.SizeHeight + 20 + b2) / 400 + 1);
                step_y = (r.SizeHeight + 20 + b2) / metk_y;
            }
            double zps = 0;
            if (r.SizeWidth < 251)
            {
                zps = 251 - r.SizeWidth;
            }
            corelApp.ActiveDocument.ActivePage.CreateLayer("метки summa");
            Layer metk_sum = corelApp.ActiveDocument.ActivePage.Layers["метки summa"];
            int gilmt = 2;
            Layer metk_g = corelApp.ActiveDocument.ActivePage.Layers.Find("метки для гильотины");
            if (metk_g != null)
            {
                gilmt = 0;
            }

            for (int q = 0; q < metk_y + 1;q ++)
            {
                Shape met = metk_sum.CreateRectangle(r.LeftX - 11 - gilmt - zps, r.BottomY - 13 + q * step_y, r.LeftX - 8 - gilmt - zps, r.BottomY - 16 + q * step_y);
                met.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                Shape met2 = metk_sum.CreateRectangle(r.RightX + 11 + gilmt, r.BottomY - 13 + q * step_y, r.RightX + 8 + gilmt, r.BottomY - 16 + q * step_y);
                met2.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met2.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }
            Shape met3 = metk_sum.CreateRectangle(r.LeftX - zps, r.BottomY - 13, r.RightX, r.BottomY - 16);
            met3.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
            met3.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            Shape met4 = null;
            if (barc2 == true)
            {
                met4 = metk_sum.CreateRectangle(r.LeftX - zps, r.TopY + 16, r.RightX, r.TopY + 13);
                met4.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met4.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }

            int n_met = metk_y + 1;
            int x_dis = (int)Math.Round(step_y / 0.025);
            int y_dis = (int)Math.Round((met3.SizeWidth + 20 + 3) / 0.025);

            Layer brk = corelApp.ActivePage.CreateLayer("баркод");
            Random rnd = new Random();
            long barnmbr = ((rnd.Next(100000000,999999999)) * 9);
            long nbr1 = long.Parse("6" + barnmbr);
            long nbr2 = long.Parse("9" + barnmbr);
            long sm1 = nbr1.ToString()
                           .Where(char.IsDigit)
                           .Sum(c => c - 0);
            long nm1 = 10 - (sm1 % 10);
            MessageBox.Show(nm1.ToString());
            long sm2 = nbr2.ToString()
                           .Where(char.IsDigit)
                           .Sum(c => c - 0);
            long nm2 = 10 - (sm2 % 10);
            Shape barcode1 = brk.CreateArtisticText(met3.RightX, met3.TopY, "S" + nbr1.ToString() + nm1.ToString() + "S", cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "POSTNET", 30, cdrTriState.cdrTrue, cdrTriState.cdrFalse, cdrFontLine.cdrNoFontLine, cdrAlignment.cdrNoAlignment);
            barcode1.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
            barcode1.ConvertToCurves();
            barcode1.SetSize(212, 9.53);
            barcode1.RightX = met3.RightX;
            barcode1.BottomY = met3.TopY;
            Shape podp1 = brk.CreateArtisticText(met3.RightX - 242,met3.TopY+3,nbr1.ToString(), cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrTrue, cdrTriState.cdrFalse, cdrFontLine.cdrNoFontLine, cdrAlignment.cdrNoAlignment);
            podp1.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
            podp1.ConvertToCurves();

            if(barc2 == true)
            {
                Shape barcode2 = brk.CreateArtisticText(met4.LeftX, met4.TopY, "S" + nbr2.ToString() + nm2.ToString() + "S", cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "POSTNET", 30, cdrTriState.cdrTrue, cdrTriState.cdrFalse, cdrFontLine.cdrNoFontLine, cdrAlignment.cdrNoAlignment);
                barcode2.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                barcode2.ConvertToCurves();
                barcode2.SetSize(212, 9.53);
                barcode2.RightX = met4.LeftX + 212;
                barcode2.TopY = met4.BottomY;
                barcode2.Rotate(180);
                Shape podp2 = brk.CreateArtisticText(met4.LeftX + 216, met4.BottomY - 6, nbr2.ToString(), cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrTrue, cdrTriState.cdrFalse, cdrFontLine.cdrNoFontLine, cdrAlignment.cdrNoAlignment);
                podp2.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                podp2.ConvertToCurves();
                podp2.Rotate(180);
            }
            corelApp.ActiveDocument.EndCommandGroup();
            
        }   
    
    }

    
}
