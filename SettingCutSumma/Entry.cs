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
using Application = Corel.Interop.VGCore.Application;
using corel = Corel.Interop.VGCore;

namespace SummaMetki
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("b3f7d9a3-02f6-4259-9b13-28c2c8070cfa")]
   
    public class Settings_cut //Параметры plt
    {
        public string path_plt { get; set; } = @"C:\РЕЗКА";
        public int velosity { get; set; } = 600;
        public bool barc2 { get; set; } = true;
        public int overcut { get; set; } = 1;
        public bool smothing { get; set; } = true;
        public string[] color_name { get; set; } = new[] { "ColorContour" };
        public bool doc_name { get; set; } = true;
    }
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("107368b5-9074-4f33-b2b6-4ce9852f503f")]
    public class Entry
    {
       

        public void Initialize(Application app)
        {
            corelApp = app;
        }



        
        private Thread uiThread;
        private MainWindow win;
        private progress win1;
        private Application corelApp;

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

                win = new MainWindow(corelApp);
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
            uiThread = new Thread(() =>
            {
                win1 = new progress();
                win1.Show();
                win1.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                System.Windows.Threading.Dispatcher.Run();
            });
            uiThread.SetApartmentState(ApartmentState.STA);
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
            
            
                try
                {
                    StartMacros();
                }
                finally
                {
                    
                    Progress_stop();
                }
            

        }
        public void StartMacros()
        {
            
            string name = corelApp.ActiveDocument.Name;
            var corelDoc = corelApp.ActiveDocument;
            var ActPage = corelDoc.ActivePage;
            var layers = ActPage.Layers;
            ShapeRange allShapes = ActPage.Shapes.All();
            corelApp.BeginDraw(false,true,false,true);
            try
            {
                corelDoc.BeginCommandGroup("Добавление меток и баркодов для " + name);
                corelDoc.Unit = corel.cdrUnit.cdrMillimeter;
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

                Layer prnt = layers.Find("печать");
                if (prnt == null)
                {
                    corelApp.ActiveLayer.Name = "печать";
                    prnt = layers["печать"];
                }

                Layer rzk = layers.Find("резка");
                if (rzk == null)
                {
                    ActPage.CreateLayer("резка");
                    rzk = layers["резка"];
                }
                ActPage.CreateLayer("метки summa");
                Layer metk_sum = layers["метки summa"];
                Layer metk_g = layers.Find("метки для гильотины");
                Layer brk = ActPage.CreateLayer("баркод");
                Layer nmDoc = null;

                //ищем контур по имени спотового цвета абриса
                foreach (string color in settings.color_name)
                {
                    ShapeRange cut = corelApp.ActiveDocument.ActivePage.Shapes.FindShapes(Query: "@outline.color.name='" + color + "'");
                    cut.MoveToLayer(rzk);
                    cut.BreakApart();
                }
                if (rzk.Shapes.Count == 0)
                {
                    Progress_stop();
                    MessageBox.Show("Контур резки не найден!");
                    corelApp.EndDraw();
                    return;
                }
                ShapeRange rzkShapes = rzk.Shapes.All();
                prnt.Shapes.All().CreateSelection();
                rzkShapes.AddToSelection();
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
                    nmDoc = ActPage.CreateLayer("Имя документа");
                    string nameDocVolume = name;
                    nameText = nmDoc.CreateParagraphText(r.LeftX, r.TopY, r.LeftX + r.SizeWidth + zps, r.BottomY, nameDocVolume, cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrFalse, cdrTriState.cdrFalse);
                    nameText.ConvertToCurves();
                    nameText.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                    allShapes.Add(nameText);
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
                    allShapes.Add(met0);
                    Shape met1 = metk_sum.CreateRectangle(r.RightX + 11 + gilmt, r.BottomY - 12.53 + q * step_y, r.RightX + 8 + gilmt, r.BottomY - 15.53 + q * step_y);
                    met1.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                    met1.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                    allShapes.Add(met1);
                }
                if (settings.doc_name == true)
                {
                    nameText.RightX = r.RightX;
                    nameText.TopY = metk_sum.Shapes.All().TopY - b2;
                }
                ShapeRange oposAndBar = corelDoc.CreateShapeRangeFromArray();
                Shape met;
                oposAndBar = Create_OPOS_XY("6");// вызов отрисовки нижней метки OPOS_XY и баркода
                int n_met = metk_y + 1;
                int x_dis = (int)Math.Round(step_y / 0.025);
                int y_dis = (int)Math.Round((oposAndBar.SizeWidth + 20 + 3) / 0.025);
                ExportCut();
                allShapes.Rotate(-270); // возвращаем исходное положение
                if (barc2 == true) //проверка нужен ли второй баркод
                {
                    oposAndBar = Create_OPOS_XY("9"); // повторно отрисовываем нижнюю метку, если выбран второй баркод
                    oposAndBar.Rotate(180);
                    oposAndBar.TopY = metk_sum.Shapes.All().TopY;
                    allShapes.Rotate(180);
                    rzkShapes = rzk.Shapes.All();
                    rzkShapes.BreakApart();
                    rzkShapes.Sort("@shape1.CenterY * 100 - @shape1.CenterX < @shape2.CenterY * 100 - @shape2.CenterX");
                    rzkShapes.Combine();
                    ExportCut();
                    allShapes.Rotate(270);
                }

                ShapeRange Create_OPOS_XY(string n)
                {
                    met = metk_sum.CreateRectangle(r.LeftX - zps, r.BottomY - 12.53, r.RightX, r.BottomY - 15.53); // отрисовка нижней метки OPOS_XY
                    met.Fill.UniformColor.CMYKAssign(0, 0, 0, 100);
                    met.Style.StringAssign(@"{""outline"":{""width"":""0""}}");
                    allShapes.Add(met);
                    oposAndBar = BarcodeEdit(n);
                    oposAndBar.Add(met);
                    allShapes.AddRange(oposAndBar);
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
                    Shape barshape = barcode_create.Create(brk, bar); // вызов отрисовки баркода
                    barshape.RightX = met.RightX;
                    barshape.BottomY = met.TopY;
                    Shape podp = brk.CreateArtisticText(met.RightX - 245, met.TopY + 3, nbr.ToString(), cdrTextLanguage.cdrEnglishUS, cdrTextCharSet.cdrCharSetDefault, "Arial", 12, cdrTriState.cdrFalse, cdrTriState.cdrFalse); // подписываем значение баркода
                    podp.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                    podp.ConvertToCurves();
                    oposAndBar = corelDoc.CreateShapeRangeFromArray();
                    oposAndBar.Add(barshape);
                    oposAndBar.Add(podp);// собираем баркод и подпись в один ShapeRange
                    return oposAndBar;
                }
                void ExportCut()
                {
                    //ненужные слои перед экспортом делаем непечатными
                    foreach (Layer layer in layers)
                    {
                        if (layer != rzk)
                        {
                            layer.Printable = false;
                        }
                    }
                    rzkShapes.Sort("@shape1.CenterY * 100 - @shape1.CenterX < @shape2.CenterY * 100 - @shape2.CenterX"); // упорядочиваем контуры резки 
                    rzkShapes.Combine();
                    allShapes.AddRange(rzk.Shapes.All());
                    var convert_plt = new Convert_to_plt_and_export();
                    convert_plt.Init(corelApp);
                    convert_plt.Open_pdf(nbr, n_met, x_dis, y_dis);

                }
                foreach (Layer layer in layers)
                {
                    if (layer != rzk)
                    {
                        layer.Printable = true;
                    }
                    rzk.Printable = false;
                }
                ActPage.Shapes.All().AlignRangeToPage(cdrAlignType.cdrAlignVCenter);
                ActPage.Shapes.All().AlignRangeToPage(cdrAlignType.cdrAlignHCenter);
                ActPage.SizeHeight = allShapes.SizeHeight + 2;
                ActPage.SizeWidth = allShapes.SizeWidth + 2;
                corelDoc.ClearSelection();
                corelDoc.EndCommandGroup();
                corelApp.EndDraw();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Во время работы плагина произошла ошибка" + ex.Message);
            }
            
        }
            


        }   
    
    }

    

