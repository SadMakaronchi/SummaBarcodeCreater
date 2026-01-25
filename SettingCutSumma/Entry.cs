using Corel.Interop.VGCore;
using SettingCutSumma;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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
        public string[]color_name {  get; set; }
        public bool doc_name {  get; set; }
    }
   
    public class Entry
    {
        public corel.Application corelApp;
        public void Init(corel.Application app)
        {
            corelApp = app;
        }
        


        public corel.Application crl = new corel.Application();
        private Thread uiThread;
        private MainWindow win;
        private progress win1;
        public void Start() //запуск окна настроек в отдельном потоке
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

                win = new MainWindow(crl);
                win.Closed += (s, e) => win.Dispatcher.InvokeShutdown();
                win.Show();
                win.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);

                System.Windows.Threading.Dispatcher.Run();
            });

            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.IsBackground = true;
            uiThread.Start();
        }

        
        public void Progress() //запуск полосы прогресса экспорта в отдельнном потоке
        {
            if (uiThread != null && uiThread.IsAlive)
            {
                win1.Dispatcher.Invoke(() =>
                {
                    if (win1.WindowState == WindowState.Minimized)
                        win1.WindowState = WindowState.Normal;
                    win1.Activate();
                });
                return;
            }
            uiThread = new Thread(() =>
            {

                win1 = new progress();
                win1.Closed += (s, e) => win1.Dispatcher.InvokeShutdown();
                win1.Show();
                win1.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);

                System.Windows.Threading.Dispatcher.Run();
            });

            uiThread.SetApartmentState(ApartmentState.STA); 
            uiThread.IsBackground = true;
            uiThread.Start();
        }



        public void Progress_stop()//остановка полосы прогресса экспорта
        {
            if (win1 != null)
            {
                win1.Dispatcher.Invoke(() =>
                {
                    win1.Close();
                    win1 = null;
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
            // получаем настройки из файла
            Settings_cut settings = new Settings_cut();
            string path_settings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
            using (FileStream fs = File.OpenRead(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                settings = (Settings_cut)xsz.Deserialize(fs);
            }
            bool barc2 = settings.barc2;
            long nbr;
            string[] colorCut = settings.color_name;

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
            Layer nmDoc = null;

            //ищем контур по имени спотового цвета абриса
            foreach (string color in colorCut)
            {
                ShapeRange cut = corelApp.ActiveDocument.ActivePage.Shapes.FindShapes(Query: "@outline.color.name='" + color + "'");
                cut.MoveToLayer(rzk);
                foreach (Shape s in cut)
                {
                    s.BreakApart();
                }
            }
            if(rzk.Shapes.Count==0)
            {
                MessageBox.Show("Контур резки не найден!");
                corelApp.EndDraw();
                Progress_stop();
                return;
            }
            prnt.Shapes.All().CreateSelection();
            rzk.Shapes.All().AddToSelection();
            ShapeRange r = corelApp.ActiveSelectionRange;
            Shape nameText = null;
            double hgtName = 0;
            double zps = 0;
            if (r.SizeWidth < 251) // проверяем ширину макета, чтобы при необходимости добавить пустое место, для того чтобы влез баркод
            {
                zps = 251 - r.SizeWidth;
            }

            if (settings.doc_name == true) //создаём название документа
            {
                nmDoc = corelApp.ActiveDocument.ActivePage.CreateLayer("Имя документа");
                string nameDocVolume = corelApp.ActiveDocument.Name;
                nameText = nmDoc.CreateParagraphText(r.LeftX,r.TopY,r.LeftX + r.SizeWidth + zps,r.BottomY, nameDocVolume, cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrFalse, cdrTriState.cdrFalse);
                nameText.ConvertToCurves();
                nameText.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                hgtName = nameText.SizeHeight + 2;
            }
            // считаем шаг и количество меток OPOS
            double b2 = 0;
            if (barc2 == true)
            {
                b2 = 15.53;
            }
            int metk_y = (int)Math.Ceiling((r.SizeHeight + 15.53 + b2 + hgtName) / 400);
            double step_y = (r.SizeHeight + 15.53 + b2 + hgtName) / metk_y;
            if (step_y > 410)
            {
                metk_y = (int)Math.Ceiling((r.SizeHeight + 15.53 + b2 + hgtName) / 400 + 1);
                step_y = (r.SizeHeight + 15.53 + b2 + hgtName) / metk_y;
            }
            

            
            
            int gilmt = 2;
            
            if (metk_g != null)
            {
                gilmt = 0;
            }

            for (int q = 0; q < metk_y + 1; q++) // отрисовка меток OPOS
            {
                Shape met0 = metk_sum.CreateRectangle(r.LeftX - 11 - gilmt - zps, r.BottomY - 12.53 + q * step_y, r.LeftX - 8 - gilmt - zps, r.BottomY - 15.53 + q * step_y);
                met0.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met0.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                Shape met1 = metk_sum.CreateRectangle(r.RightX + 11 + gilmt, r.BottomY - 12.53 + q * step_y, r.RightX + 8 + gilmt, r.BottomY - 15.53 + q * step_y);
                met1.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                met1.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
            }
            if (settings.doc_name == true)
            {
                nameText.RightX = r.RightX;
                nameText.TopY = metk_sum.Shapes.All().TopY - b2;
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
                Shape podp = brk.CreateArtisticText(met.RightX - 245, met.TopY + 3, nbr.ToString(), cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrFalse, cdrTriState.cdrFalse); // подписываем значение баркода
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
                if (settings.doc_name == true)
                {
                    nmDoc.Printable = false;
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
            if(settings.doc_name == true)
            {
                nmDoc.Printable = true;
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

    

