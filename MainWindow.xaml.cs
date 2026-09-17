using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Xml.Serialization;
using Application = Corel.Interop.VGCore.Application;
using corel = Corel.Interop.VGCore;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace SummaBarcodeCreater
{
   
    public partial class MainWindow : Window
    {
        public corel.Application corelApp;
        public string folder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel");
        public string path_settings = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");
        public string fn;
        public Settings settings;
        public Settings_define settings_define;
        public CommonOpenFileDialog dialog;
        public StylesController stylesController;
        public class ComboItems
        {
            public int Value {  get; set; }
            public string pref {  get; set; }
            public override string ToString() => pref;
        }
        public MainWindow(Application app, CommonOpenFileDialog dialog)
        {
            corelApp = app;
            settings_define = new Settings_define();
            settings = settings_define.Get_settings();
            this.dialog = dialog;
            InitializeComponent();
            try
            {
                stylesController = new StylesController(Resources,corelApp);
            }
            catch
            {
               MessageBox.Show("VGCore Erro");
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
            PlxCheck.IsChecked = settings.plx;
            Color.Text = string.Join(",", settings.color_name);
            Color.IsMouseDirectlyOverChanged += BleedChanged;
            NameDoc.IsChecked = settings.doc_name;
            UpBleed.Text = settings.up_bleed.ToString() + " мм";
            DownBleed.Text = settings.down_bleed.ToString() + " мм";
            LeftBleed.Text = settings.left_bleed.ToString() + " мм";
            RightBleed.Text = settings.right_bleed.ToString() + " мм";
            UpBleed.IsMouseDirectlyOverChanged += BleedChanged;
            UpBleed.PreviewTextInput += Bleed_TextChanged;
            DownBleed.IsMouseDirectlyOverChanged += BleedChanged;
            DownBleed.PreviewTextInput += Bleed_TextChanged;
            LeftBleed.IsMouseDirectlyOverChanged += BleedChanged;
            LeftBleed.PreviewTextInput += Bleed_TextChanged;
            RightBleed.IsMouseDirectlyOverChanged += BleedChanged;
            RightBleed.PreviewTextInput += Bleed_TextChanged;
        }

        private void Bleed_TextChanged(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !e.Text.All(char.IsDigit);
        }

        
        private void BleedChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if(UpBleed.IsFocused == true)
            {
                UpBleed.Text = "";
            }
            else
            {
                if (UpBleed.Text == "")
                {
                    UpBleed.Text = settings.up_bleed.ToString() + " мм";
                }
                else
                {
                    if (UpBleed.Text.Contains(" мм") == false)
                    {
                        UpBleed.Text = UpBleed.Text + " мм";
                    }
                }
            }
            if (DownBleed.IsFocused == true)
            {
                DownBleed.Text = "";
            }
            else
            {
                if (DownBleed.Text == "")
                {
                    DownBleed.Text = settings.down_bleed.ToString() + " мм";
                }
                else
                {
                    if (DownBleed.Text.Contains(" мм") == false)
                    {
                        DownBleed.Text = DownBleed.Text + " мм";
                    }
                }
            }
            if (LeftBleed.IsFocused == true)
            {
                LeftBleed.Text = "";
            }
            else
            {
                if (LeftBleed.Text == "")
                {
                    LeftBleed.Text = settings.left_bleed.ToString() + " мм";
                }
                else
                {
                    if (LeftBleed.Text.Contains(" мм") == false)
                    {
                        LeftBleed.Text = LeftBleed.Text + " мм";
                    }
                }
            }
            if (RightBleed.IsFocused == true)
            {
                RightBleed.Text = "";
            }
            else
            {
                if (RightBleed.Text == "")
                {
                    RightBleed.Text = settings.right_bleed.ToString() + " мм";
                }
                else
                {
                    if (RightBleed.Text.Contains(" мм") == false)
                    {
                        RightBleed.Text = RightBleed.Text + " мм";
                    }
                }
            }

        }
        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();

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
            settings.plx = PlxCheck.IsChecked == true;
            if (UpBleed.Text.Contains(" мм") == true)
            {
                settings.up_bleed = double.Parse(UpBleed.Text.Substring(0, UpBleed.Text.Length - 3));
            }
            else
            {
                settings.up_bleed = double.Parse(UpBleed.Text);
            }
            if (DownBleed.Text.Contains(" мм") == true)
            {
                settings.down_bleed = double.Parse(DownBleed.Text.Substring(0, DownBleed.Text.Length - 3));
            }
            else
            {
                settings.down_bleed = double.Parse(DownBleed.Text);
            }
            if (LeftBleed.Text.Contains(" мм") == true)
            {
                settings.left_bleed = double.Parse(LeftBleed.Text.Substring(0, LeftBleed.Text.Length - 3));
            }
            else
            {
                settings.left_bleed = double.Parse(LeftBleed.Text);
            }
            if (RightBleed.Text.Contains(" мм") == true)
            {
                settings.right_bleed = double.Parse(RightBleed.Text.Substring(0, RightBleed.Text.Length - 3));
            }
            else
            {
                settings.right_bleed = double.Parse(RightBleed.Text);
            }
            //пишем в файл
            using (FileStream fs = File.Create(path_settings))
            {
                XmlSerializer xsz = new XmlSerializer(typeof(Settings));
                xsz.Serialize(fs,settings);
            }
            MessageBox.Show("Настройки резки сохранены!");
            this.Close();
        }

        [STAThread]
        public void FolderClick(object sender, RoutedEventArgs e)
        {
           
            if (dialog.ShowDialog(this) == CommonFileDialogResult.Ok)
            {
                fn = dialog.FileName;
                Path.Text = fn;
            }


        }

    }
    

}
