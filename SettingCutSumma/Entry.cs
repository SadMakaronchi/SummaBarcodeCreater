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
    public class Entry
    {
        public corel.Application corelApp;
        public void Init(corel.Application app)
        {
            corelApp = app;
        }
        public class Settings_cut
        {
            public string path_plt { get; set; } = @"C:\РЕЗКА\";
            public int velosity { get; set; } = 600;
            public bool barc2 { get; set; } = true;
            public double overcut { get; set; } = 0.1;       
        }


        public corel.Application crl = new corel.Application();
        public void Start()
        {
            
            
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"SummaPanel", "SettingsCutSumma.xml");
                using (FileStream fs = File.Create(path))
                {
                    XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                    var settings = new Settings_cut();
                    xsz.Serialize(fs,settings);
                }
            

            
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

        public void init()
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

            for (int q = 0; q < metk_y;)
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
            if (barc2 == true)
            {
                Shape met4 = metk_sum.CreateRectangle(r.LeftX - zps, r.TopY + 16, r.RightX, r.TopY + 13);
                met4.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met4.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }

            int n_met = metk_y + 1;
            int x_dis = (int)Math.Round(step_y / 0.025);
            int y_dis = (int)Math.Round((met3.SizeWidth + 20 + 3) / 0.025);

            Layer brk = corelApp.ActivePage.CreateLayer("баркод");
            Random rnd = new Random();
            long barnmbr = ((rnd.Next(100000000,999999999)) * 9);
            string nbr1 = "6" + barnmbr;
            string nbr2 = "9" + barnmbr;
            long sm1 = nbr1.Where(char.IsDigit)
                          .Sum(c => c - 0);
            long sm2 = nbr2.Where(char.IsDigit)
                          .Sum(c => c - 0);
        }   
    
    }

    
}
