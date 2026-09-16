using Corel.Interop.VGCore;
using System;
using System.Windows.Forms;
using Application = Corel.Interop.VGCore.Application;
using corel = Corel.Interop.VGCore;
using System.IO;

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
            define_per.allShapes = define_per.ActPage.Shapes.All();
            var progress = new Progress_bar(); //запуск прогресс-бара
            progress.Progress(app);
            app.Optimization = true;
            try
            {
                define_per.corelDoc.BeginCommandGroup("Добавление меток и баркодов для " + define_per.DocName);
                
                define_per.corelDoc.Unit = corel.cdrUnit.cdrMillimeter;
                //тут создаём рабочие слои
                var LayerAndContour = new LayerAndSearchContour();
                LayerAndSearchContour layerAndSearchContour = LayerAndContour.CreateLayer(app,settings,define_per);
                var OPOS_marks = new OPOS_marks();
                OPOS_marks.Create_marks(layerAndSearchContour,settings,define_per);//метки OPOS
                var numbers = new Numbers();//числа баркода
                numbers.Calc();
                var create_barcode = new CreateBarcode();
                create_barcode.Create(layerAndSearchContour, numbers, settings,define_per);
                var export_plt = new ExportPlt();
                export_plt.define_obj(layerAndSearchContour,settings,define_per,numbers,OPOS_marks);
                define_per.allShapes.AlignRangeToPage(cdrAlignType.cdrAlignVCenter);
                define_per.allShapes.AlignRangeToPage(cdrAlignType.cdrAlignHCenter);
                define_per.ActPage.SizeHeight = define_per.allShapes.SizeHeight + settings.up_bleed + settings.down_bleed;
                define_per.ActPage.SizeWidth = define_per.allShapes.SizeWidth + settings.right_bleed + settings.left_bleed;
                layerAndSearchContour.rzk.Printable = false;
                layerAndSearchContour.rzk.Visible = false;
                if (settings.plx == true)
                {
                    bool rotate = false;
                    File.WriteAllText((Path.Combine(settings.path_plt, numbers.down_number.ToString() + ".plx")), createPlx.Plx(define_per, LayerAndContour, settings,rotate));
                }
                if (settings.plx == true&&settings.barc2 == true)
                {
                    bool rotate = true;
                    File.WriteAllText((Path.Combine(settings.path_plt, numbers.up_number.ToString() + ".plx")), createPlx.Plx(define_per, LayerAndContour, settings,rotate));
                }
                app.Optimization = false;
                app.Refresh();
                progress.Progress_stop();
                define_per.corelDoc.EndCommandGroup();
            }
            catch (NullReferenceException ex)
            {
                
                progress.Progress_stop();
                app.Optimization = false;
                define_per.corelDoc.EndCommandGroup();
                MessageBox.Show("Во время работы плагина произошла ошибка" + ex);
            }

        }
    }
}
