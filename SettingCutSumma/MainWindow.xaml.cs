using Corel.Interop.VGCore;
using System;
using System.Text;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using corel = Corel.Interop.VGCore;
using Window = System.Windows.Window;
using MessageBox = System.Windows.MessageBox;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Linq;

namespace SummaMetki
{
   
    public partial class MainWindow : Window
    {
       
        public corel.Application corelApp;
        public Styles.StylesController stylesController;
        Settings_cut settings = new Settings_cut();
        public string path_settings = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
        public string fn;
        public class ComboItems
        {
            public int Value {  get; set; }
            public string pref {  get; set; }
            public override string ToString() => pref;
        }
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
                Velosity.Items.Add(new ComboItems
                {
                    Value = vel,
                    pref = vel + " мм/сек"
                });
            }
            for (int ovr = 1; ovr < 11; ovr++)
            {
                Overcut.Items.Add(new ComboItems
                {
                    Value = ovr,
                    pref = ovr / 10m + "мм"
                });
            }
            //подставляем значения переменных из файла
            Velosity.SelectedItem = Velosity.Items.Cast<ComboItems>().First(x => x.Value == settings.velosity);
            Overcut.SelectedItem = Overcut.Items.Cast<ComboItems>().First(x => x.Value == settings.overcut);
            Smothing.IsChecked = settings.smothing;
            Barcode2.IsChecked = settings.barc2;
            fn = settings.path_plt;
            Path.Text = fn;
            Color.Text = string.Join(",", settings.color_name);
            NameDoc.IsChecked = settings.doc_name;
        }
        public void Ok_click(object sender, RoutedEventArgs e)
        {
            //обновляем переменнные
            settings.velosity = (int)((ComboItems)Velosity.SelectedItem).Value; ;
            settings.overcut = (int)((decimal)((ComboItems)Overcut.SelectedItem).Value);
            settings.smothing = Smothing.IsChecked == true;
            settings.barc2 = Barcode2.IsChecked == true;
            settings.path_plt = fn;
            settings.color_name = Color.Text.Split(new char[] {','});
            settings.doc_name = NameDoc.IsChecked == true;
           
            //пишем в файл
            using (FileStream fs = File.Create(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                xsz.Serialize(fs,settings);
            }
            MessageBox.Show("Настройки резки сохранены!");
            this.Close();
        }
        
        public void FolderClick(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog() { IsFolderPicker = true, DefaultDirectory = fn, Title = "Выберите папку для сохранения файла резки в формате plt" };
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
