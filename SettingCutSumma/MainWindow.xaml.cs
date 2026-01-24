using Corel.Interop.VGCore;
using System;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
using corel = Corel.Interop.VGCore;
using Window = System.Windows.Window;
using MessageBox = System.Windows.MessageBox;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace SummaMetki
{
   
    public partial class MainWindow : Window
    {
       
        public corel.Application corelApp;
        public Styles.StylesController stylesController;
        Settings_cut settings = new Settings_cut();
        public string path_settings = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
        public string fn;
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
            Smothing.IsChecked = settings.smothing;
            Barcode2.IsChecked = settings.barc2;
            fn = settings.path_plt;
            Path.Text = fn;
        }
        public void Ok_click(object sender, RoutedEventArgs e)
        {
            //обновляем переменнные
            settings.velosity = (int)Velosity.SelectedItem;
            settings.overcut = (int)((decimal)Overcut.SelectedItem * 10);
            settings.smothing = Smothing.IsChecked == true;
            settings.barc2 = Barcode2.IsChecked == true;
            settings.path_plt = fn;
            //пишем в файл
            using (FileStream fs = File.Create(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                xsz.Serialize(fs,settings);
            }
            MessageBox.Show("Настройки резки сохранены!");
            this.Close();
        }
        public void Smoth_click(object sender, RoutedEventArgs e)
        {
           
        }
        public void BarcClick(object sender, RoutedEventArgs e)
        {
            
        }
        public void VelChange(object sender, RoutedEventArgs e)
        {
           
        }
        public void OverChange(object sender, RoutedEventArgs e)
        {
            
        }
        public void FolderClick(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog() { IsFolderPicker = true };
            dialog.Title = "Выберите папку для сохранения файла резки в формате plt";
            dialog.DefaultDirectory = fn;
            if(dialog.ShowDialog()== CommonFileDialogResult.Ok)
            {
                fn = dialog.FileName;
                Path.Text = fn;
            }
        }
            
           

       

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }
    }
    

}
