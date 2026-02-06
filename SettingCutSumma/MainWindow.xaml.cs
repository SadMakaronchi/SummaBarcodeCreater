using Microsoft.WindowsAPICodePack.Dialogs;
using SettingCutSumma;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Shapes;
using System.Xml.Serialization;
using corel = Corel.Interop.VGCore;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace SummaMetki
{
   
    public partial class MainWindow : Window
    {
       
        public corel.Application corelApp;
        public Styles.StylesController stylesController;
        Settings_cut settings = new Settings_cut();
        public string folder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel");
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
            try
            {
                
                try
                {
                    Directory.CreateDirectory(folder);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Нет прав на создание папки для хранения файла с настройками!Пожалуйста запустите CorelDraw с правами администратора");
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Ошибка создания папки для файлов с настройками:" + ex.Message);
                }
                if (File.Exists(path_settings) == true)
                {
                    using (FileStream fs = File.OpenRead(path_settings))
                    {
                        XmlSerializer xsz = new XmlSerializer(typeof(Settings_cut));
                        settings = (Settings_cut)xsz.Deserialize(fs);
                    }
                }
                else
                {
                    MessageBox.Show("Не найден файл настроек,настройки будут заданы по умолчанию");
                }
            }
            catch
            {
               MessageBox.Show("Файл настроек повреждён,будут заданы настройки по умолчанию");
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
            Path.Text = settings.path_plt;
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
            
            if (Directory.Exists(Path.Text) == false)
            {
                MessageBox.Show("Отсутствует папка экспорта plt по заданному пути");
                Path.Text = @"C:\";
            }
            
            var dialog = new CommonOpenFileDialog() { IsFolderPicker = true, InitialDirectory = Path.Text, Title = "Выберите папку для сохранения файла резки в формате plt" };
                
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
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
