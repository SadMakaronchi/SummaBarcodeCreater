using Corel.Interop.VGCore;
using System;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
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
            Settings_cut settings = new Settings_cut();
            string path_settings = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
            using (FileStream fs = File.OpenRead(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                settings = (Settings_cut)xsz.Deserialize(fs);
            }

            for (int vel = 100; vel < 1010; vel += 10)
            {
                Velosity.Items.Add(vel);
            }
            for(int ovr = 1 ; ovr < 11; ovr++)
            {
                Overcut.Items.Add(ovr/10m);
            }

            Velosity.SelectedItem = settings.velosity;
            Overcut.SelectedItem = settings.overcut/10m;
              

                 
           
        }

            
           

       

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }
    }
    

}
