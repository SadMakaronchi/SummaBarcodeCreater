using Corel.Interop.VGCore;
using System.Windows.Forms;
using Application = Corel.Interop.VGCore.Application;
using corel = Corel.Interop.VGCore;

namespace SummaBarcodeCreater
{
    public class LayerAndSearchContour
    {
        public Application corelApp;
        public Layer prnt;
        public Layer rzk;
        public Layer metk_sum;
        public Layer metk_g;
        public Layer brk;
        public ShapeRange r;
        public ShapeRange kisscut;
        public Settings settings;
        public Define_per define_per;
        public LayerAndSearchContour CreateLayer(Application app,Settings settings,Define_per define_per)
        {
                this.settings = settings;
                this.define_per = define_per;
                this.corelApp = app;
                this.r = define_per.allShapes;
                string[] color_name = settings.color_name;
                var layers = define_per.ActPage.Layers;
                define_per.corelDoc.Unit = corel.cdrUnit.cdrMillimeter;

                prnt = GetOrCreateLayer("печать");
                rzk = GetOrCreateLayer("резка");
                metk_sum = GetOrCreateLayer("метки summa");
                brk = GetOrCreateLayer("баркод");

                define_per.allShapes.MoveToLayer(prnt);
                //ищем контур по имени спотового цвета абриса
                foreach (string color in color_name)
                {
                    kisscut = define_per.ActPage.Shapes.FindShapes(
                                         Query: "@outline.color.name='" + color + "'");
                    kisscut.MoveToLayer(rzk);

                }
                if (rzk.Shapes.Count == 0)
                {
                    var progress = new Progress_bar();
                    progress.Progress_stop();
                    MessageBox.Show("Контур резки не найден!");
                    define_per.corelDoc.EndCommandGroup();
                    return null;
                }
                else
                {
                    //добавляем слой с именем документа, если включена такая настройка
                    if (settings.doc_name == true) //создаём название документа
                    {
                        Layer nmDoc = define_per.ActPage.CreateLayer("Имя документа");
                        Shape nameText = nmDoc.CreateParagraphText(
                            0,
                            12,
                            r.SizeWidth,
                            0,
                              define_per.DocName,
                              cdrTextLanguage.cdrEnglishUS,
                              cdrTextCharSet.cdrCharSetDefault,
                              "Arial",
                              12,
                              cdrTriState.cdrFalse, cdrTriState.cdrFalse);
                        nameText.ConvertToCurves();
                        nameText.Fill.UniformColor.CMYKAssign(0, 0, 0, 40);
                        nameText.RightX = r.RightX;
                        nameText.BottomY = r.TopY + 2;
                        r.Add(nameText);
                    }
                return this;
                }


        }
        private Layer GetOrCreateLayer(string name)
        {
            Layer layer = define_per.ActPage.Layers.Find(name);

            if (layer == null)
            layer = define_per.ActPage.CreateLayer(name);
            return layer;
        }
    }
}
