using System.Windows;
using Application = Corel.Interop.VGCore.Application;
namespace SummaBarcodeCreater
{
    public partial class Progress : Window
    {
        public StylesController stylesController;
        public Progress(Application app)
        {
            InitializeComponent();
            try
            {
                stylesController = new StylesController(Resources,app);
            }
            catch
            {
                
            }

        }
        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();

        }


    }
}

