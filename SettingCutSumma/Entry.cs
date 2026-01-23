using Corel.Interop.VGCore;
using SettingCutSumma;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using corel = Corel.Interop.VGCore;

namespace SummaMetki
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("b3f7d9a3-02f6-4259-9b13-28c2c8070cfa")]

    public class Settings_cut //Параметры plt
    {
        public string path_plt { get; set; }
        public int velosity { get; set; } = 0;
        public bool barc2 { get; set; }
        public int overcut { get; set; }
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
        public void Progress() //запуск полосы прогресса экспорта
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
                win.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);

                System.Windows.Threading.Dispatcher.Run();
            });

            uiThread.SetApartmentState(ApartmentState.STA); 
            uiThread.IsBackground = true;
            uiThread.Start();
        }



        public void Progress_stop()//остановка полосы прогресса экспорта
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

        public void Begin() //запускаем основной код плагина
        {
            if (corelApp == null)
                throw new InvalidOperationException("Corel не инициализирован");
            if (corelApp.ActiveDocument == null)
                return;

            Progress();
            Thread corelThread = new Thread(() =>
            {
                try
                {
                    StartMacros();
                }
                finally
                {
                    // Закрываем прогресс в UI потоке
                    Progress_stop();
                }
            });

            corelThread.SetApartmentState(ApartmentState.STA);
            corelThread.Start();
        }
        public void StartMacros()
        {
            corelApp.BeginDraw();
            string name = corelApp.ActiveDocument.Name;
            corelApp.ActiveDocument.BeginCommandGroup("Добавление меток и баркодов для " + name);
            corelApp.ActiveDocument.Unit = corel.cdrUnit.cdrMillimeter;
            Boolean barc2 = true;
            long nbr;
            Random rnd = new Random();// генерируем рандомное число
            long barnmbr = rnd.Next(1_000_000, 10_000_000) * 1000L + rnd.Next(0, 1000);



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
            corelApp.ActiveDocument.ActivePage.CreateLayer("метки summa");
            Layer metk_sum = corelApp.ActiveDocument.ActivePage.Layers["метки summa"];
            Layer metk_g = corelApp.ActiveDocument.ActivePage.Layers.Find("метки для гильотины");
            Layer brk = corelApp.ActivePage.CreateLayer("баркод");


            //ищем контур по имени спотового цвета абриса
            ShapeRange cut = corelApp.ActiveDocument.ActivePage.Shapes.FindShapes(Query: "@outline.color.name='CutContour'");
            cut.MoveToLayer(rzk);
            foreach (Shape s in cut)
            {
                s.BreakApart();
            }
            

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
            
            int gilmt = 2;
            
            if (metk_g != null)
            {
                gilmt = 0;
            }

            for (int q = 0; q < metk_y + 1; q++) // отрисовка меток OPOS
            {
                Shape met0 = metk_sum.CreateRectangle(r.LeftX - 11 - gilmt - zps, r.BottomY - 13 + q * step_y, r.LeftX - 8 - gilmt - zps, r.BottomY - 16 + q * step_y);
                met0.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met0.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                Shape met1 = metk_sum.CreateRectangle(r.RightX + 11 + gilmt, r.BottomY - 13 + q * step_y, r.RightX + 8 + gilmt, r.BottomY - 16 + q * step_y);
                met1.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met1.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }
            ShapeRange oposAndBar = corelApp.ActiveDocument.CreateShapeRangeFromArray();
            Shape met;
            oposAndBar = Create_OPOS_XY("6");// вызов отрисовки нижней метки OPOS_XY и баркода
            int n_met = metk_y + 1;
            int x_dis = (int)Math.Round(step_y / 0.025);
            int y_dis = (int)Math.Round((oposAndBar.SizeWidth + 20 + 3) / 0.025);
            ExportCut();
            corelApp.ActiveDocument.ActivePage.Shapes.All().Rotate(-270); // возвращаем исходное положение
            if (barc2 == true) //проверка нужен ли второй баркод
                {
                oposAndBar = Create_OPOS_XY("9"); // повторно отрисовываем нижнюю метку, если выбран второй баркод
                oposAndBar.Rotate(180);
                oposAndBar.TopY = metk_sum.Shapes.All().TopY;
                corelApp.ActiveDocument.ActivePage.Shapes.All().Rotate(180);
                rzk.Shapes.All().BreakApart();
                rzk.Shapes.All().Sort("@shape1.CenterY * 100 - @shape1.CenterX < @shape2.CenterY * 100 - @shape2.CenterX");
                rzk.Shapes.All().Combine();
                rzk.Shapes.All().BreakApart();
                ExportCut();
                corelApp.ActiveDocument.ActivePage.Shapes.All().Rotate(270);
                }
            
            ShapeRange Create_OPOS_XY(string n)
            {
                met = metk_sum.CreateRectangle(r.LeftX - zps, r.BottomY - 13, r.RightX, r.BottomY - 16); // отрисовка нижней метки OPOS_XY
                met.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                oposAndBar = BarcodeEdit(n);
                oposAndBar.Add(met);
                return oposAndBar;
            }
            
            ShapeRange BarcodeEdit(string n) // создаём штрихкод
            {
                nbr = long.Parse(n + barnmbr); // имя plt файла и подпись

                int CalcCheckDigit(long number)
                {
                    int sum = number.ToString()
                                   .Where(char.IsDigit)
                                   .Sum(c => c - '0');
                    return (10 - (sum % 10)) % 10;
                }
                int nm = CalcCheckDigit(nbr);
                string bar = nbr.ToString() + nm; // число баркода с проверочной цифрой
                var barcode_create = new BarcodeCreater();
                barcode_create.Init(corelApp);
                ShapeRange barshape = barcode_create.Create(brk, bar); // вызов отрисовки баркода
                barshape.RightX = met.RightX;
                barshape.BottomY = met.TopY;
                Shape podp = brk.CreateArtisticText(met.RightX - 245, met.TopY + 3, nbr.ToString(), cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrTrue, cdrTriState.cdrFalse, cdrFontLine.cdrNoFontLine, cdrAlignment.cdrNoAlignment); // подписываем значение баркода
                podp.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                podp.ConvertToCurves();
                oposAndBar = corelApp.ActiveDocument.CreateShapeRangeFromArray();
                oposAndBar.AddRange(barshape);
                oposAndBar.Add(podp);// собираем баркод и подпись в один ShapeRange
                return oposAndBar;
            }
            void ExportCut()
            { 
                metk_sum.Printable = false; //ненужные слои перед экспортом делаем непечатными
                prnt.Printable = false;
                brk.Printable = false;
                if (metk_g != null)
                    {
                        metk_g.Printable = false;
                    }
                ShapeRange cut_next = rzk.Shapes.All();
                cut_next.Sort("@shape1.CenterY * 100 - @shape1.CenterX < @shape2.CenterY * 100 - @shape2.CenterX"); // упорядочиваем контуры резки 
                Shape cutfnsh = cut_next.Combine();
                cutfnsh.BreakApart();

                var convert_plt = new Convert_to_plt_and_export();
                convert_plt.Init(corelApp);
                convert_plt.Open_pdf(nbr, n_met, x_dis, y_dis);
                
            }
            metk_sum.Printable = true;
            prnt.Printable = true;
            brk.Printable = true;
            rzk.Printable = false;
            if (metk_g != null)
            {
                metk_g.Printable = true;
            }
            corelApp.ActiveDocument.ActivePage.Shapes.All().AlignRangeToPage(cdrAlignType.cdrAlignVCenter);
            corelApp.ActiveDocument.ActivePage.Shapes.All().AlignRangeToPage(cdrAlignType.cdrAlignHCenter);
            corelApp.ActiveDocument.ActivePage.SizeHeight = corelApp.ActiveDocument.ActivePage.Shapes.All().SizeHeight + 2;
            corelApp.ActiveDocument.ActivePage.SizeWidth = corelApp.ActiveDocument.ActivePage.Shapes.All().SizeWidth + 2;

            corelApp.EndDraw();
            corelApp.ActiveDocument.EndCommandGroup();
        }
            


        }   
    
    }

    

