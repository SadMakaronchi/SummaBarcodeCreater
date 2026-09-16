using System.Windows;
using System.Windows.Controls;
using Application = Corel.Interop.VGCore.Application;
namespace SummaBarcodeCreater
{
    
    /// <summary>
    /// Логика взаимодействия для ControlUI.xaml
    /// </summary>
    public partial class ControlUI : UserControl
    {
        public StylesController stylesController;
        private Application corelApp;
        public ControlUI(object app)
        {
            
            InitializeComponent();
            try
            {
                this.corelApp = app as Application;
                this.corelApp.OnApplicationEvent += CorelApp_OnApplicationEvent;
                stylesController = new StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }


        }
        private void CorelApp_OnApplicationEvent(string EventName, ref object[] Parameters)
        {
            if (EventName.Equals("WorkspaceChanged") || EventName.Equals("OnColorSchemeChanged"))
            {
                var st = new StylesController(this.Resources,this.corelApp);
                st.LoadThemeFromPreference();
            }
        }

        public void StartMacros(object sender,RoutedEventArgs e)
        {
            var entry = new Entry();
            entry.Begin(corelApp);
        }

        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();

        }

    }
}
