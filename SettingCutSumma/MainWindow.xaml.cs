using Corel.Interop.VGCore;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml.Linq;
using corel = Corel.Interop.VGCore;
using Window = System.Windows.Window;


namespace SummaMetki
{
   
    public partial class MainWindow : Window
    {
        public corel.Application corelApp;
        public Styles.StylesController stylesController;
        public MainWindow(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }
            if (corelApp.ActiveDocument != null)
            {
                update_preview();
            }

            
            corelApp.SelectionChange += CorelApp_SelectionChange;
            void CorelApp_SelectionChange()
            {
                update_preview();
            }
            corelApp.ActiveDocument.ShapeChange += ActiveDocument_ShapeChange;
            void ActiveDocument_ShapeChange(corel.Shape Shape, cdrShapeChangeScope Scope)
            {
                update_preview();
            }
           


                void update_preview()
                {
                    Dispatcher.Invoke(() =>
                    {
                        ShapeRange sel = corelApp.ActiveSelectionRange;

                        if (sel == null || sel.Count == 0)
                        {
                            preview.Source = null;
                        }

                        else
                        {
                            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", $"preview_{Guid.NewGuid()}.png");
                            ExportFilter exbtmp = corelApp.ActiveDocument.ExportBitmap(path, cdrFilter.cdrPNG, cdrExportRange.cdrSelection, cdrImageType.cdrRGBColorImage);
                            exbtmp.Finish();
                            var bmp = new System.Windows.Media.Imaging.BitmapImage();
                            bmp.BeginInit();
                            bmp.UriSource = new Uri(path);
                            bmp.CacheOption = BitmapCacheOption.OnLoad;
                            bmp.EndInit();
                            bmp.Freeze();
                            preview.Source = bmp;
                            File.Delete(path);

                        }

                    });
                }
                this.Closed += (s, e) =>
                {
                    corelApp.SelectionChange -= CorelApp_SelectionChange;
                    corelApp.ActiveDocument.ShapeChange -= ActiveDocument_ShapeChange;
                   
                };
            }

       

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }
    }
    

}
