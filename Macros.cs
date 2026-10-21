using Corel.Interop.VGCore;
using System;
using System.IO;
using System.Windows.Forms;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Define_per
    {
        public Application corelApp;
        public Document corelDoc;
        public Page ActPage;
        public string DocName;
        public ShapeRange allShapes;
        public CreateAndExportColorGateDate createPlx;
    }
    public class Macros
    {
        public void StartMacros(Application app)  
        {
            if (app.Documents.Count == 0)
            {
                DialogResult dialogResult = MessageBox.Show("Нет ни одного открытого документа!",
                                                            "Добавление меток и экспорт plt",
                                                            MessageBoxButtons.OK,
                                                            MessageBoxIcon.Warning);
            }
            else
            {
                Settings_define settings_define = new Settings_define();
                Settings settings = settings_define.Get_settings();
                CreateAndExportColorGateDate createPlx = new CreateAndExportColorGateDate();
                var define_per = new Define_per();
                define_per.corelApp = app;
                define_per.corelDoc = app.ActiveDocument;
                define_per.ActPage = app.ActiveDocument.ActivePage;
                

                if (app.ActiveDocument.Name.Substring(app.ActiveDocument.Name.Length - 4, 4) == ".cdr")
                {
                    define_per.DocName = app.ActiveDocument.Name.Substring(0, app.ActiveDocument.Name.Length - 4);
                }
                else
                {
                    define_per.DocName = app.ActiveDocument.Name;
                }
                define_per.allShapes = define_per.corelDoc.CreateShapeRangeFromArray();
                foreach (Shape s in define_per.ActPage.Shapes.All())
                {
                    if (s.Type!=cdrShapeType.cdrGuidelineShape)
                    {
                        define_per.allShapes.Add(s);
                    }
                }
                var progress = new Progress_bar(); //запуск прогресс-бара
                progress.Progress(app);
                app.Optimization = true;

                try
                {
                    define_per.corelDoc.BeginCommandGroup("Добавление меток и баркодов для " + define_per.DocName);

                    define_per.corelDoc.Unit = cdrUnit.cdrMillimeter;
                    //тут создаём рабочие слои
                    var LayerAndContour = new LayerAndSearchContour();
                    LayerAndSearchContour layerAndSearchContour = LayerAndContour.CreateLayer(app, settings, define_per);//создаём слои и распределяем по слоям
                    if (layerAndSearchContour == null)
                    {
                        progress.ShowMessage("Контур резки не найден!");
                        return;
                    }
                    var OPOS_marks = new OPOS_marks();
                    OPOS_marks.Create_marks(layerAndSearchContour, settings, define_per);//метки OPOS
                    var numbers = new Numbers();//числа баркода
                    numbers.Calc();
                    var create_barcode = new CreateBarcode();
                    create_barcode.Create(layerAndSearchContour, numbers, settings, define_per);//создание баркодов
                    var pdf = new ExportToPDF();
                    pdf.BeginExport(define_per.corelApp, numbers, settings, define_per, LayerAndContour);
                    //var export_plt = new ExportPlt();//экспорт
                    //export_plt.define_obj(layerAndSearchContour, settings, define_per, numbers, OPOS_marks);
                    //define_per.allShapes.AlignRangeToPage(cdrAlignType.cdrAlignVCenter);
                    //define_per.allShapes.AlignRangeToPage(cdrAlignType.cdrAlignHCenter);
                    //layerAndSearchContour.rzk.Printable = false;
                    //layerAndSearchContour.rzk.Visible = false;
                    //if (settings.plx == true)
                    //{
                    //    define_per.ActPage.SizeHeight = define_per.allShapes.SizeHeight;
                    //    define_per.ActPage.SizeWidth = define_per.allShapes.SizeWidth;
                    //    layerAndSearchContour.metk_sum.Printable = false;
                    //    layerAndSearchContour.brk.Printable = false;
                    //    if (layerAndSearchContour.nmDoc != null)
                    //    {
                    //        layerAndSearchContour.nmDoc.Printable = false;
                    //    }
                    //    bool rotate = false;
                    //    File.WriteAllText((Path.Combine(settings.path_plt, numbers.down_number.ToString() + ".plx")),
                    //                                    createPlx.Plx(define_per, 
                    //                                                  LayerAndContour, 
                    //                                                  settings, 
                    //                                                  rotate,
                    //                                                  export_plt.plt1));
                    //}
                    //if (settings.plx == true && settings.barc2 == true)
                    //{
                    //    bool rotate = true;
                    //    File.WriteAllText((Path.Combine(settings.path_plt, numbers.up_number.ToString() + ".plx")), 
                    //                                    createPlx.Plx(define_per, 
                    //                                    LayerAndContour, 
                    //                                    settings, 
                    //                                    rotate,
                    //                                    export_plt.plt2));
                    //}
                    
                    
                        layerAndSearchContour.metk_sum.Printable = true;
                        layerAndSearchContour.brk.Printable = true;
                        if (layerAndSearchContour.nmDoc != null)
                        {
                            layerAndSearchContour.nmDoc.Printable = true;
                        }
                    
                    define_per.ActPage.SizeHeight = define_per.allShapes.SizeHeight 
                                                  + settings.up_bleed 
                                                  + settings.down_bleed;
                    define_per.ActPage.SizeWidth = define_per.allShapes.SizeWidth 
                                                 + settings.left_bleed 
                                                 + settings.right_bleed;
                    app.Optimization = false;
                    app.Refresh();
                }
                catch (Exception ex)
                {
                    progress.ShowMessage("Во время работы плагина произошла ошибка" + ex);
                }
                finally
                {
                    if (progress.win != null)
                    { 
                        progress.Progress_stop();
                    }
                    app.Optimization = false;
                    define_per.corelDoc.EndCommandGroup();
                }
            }
        }
    }
}
