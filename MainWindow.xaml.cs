using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Serialization;
using Application = Corel.Interop.VGCore.Application;
using ComboBox = System.Windows.Controls.ComboBox;
using corel = Corel.Interop.VGCore;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
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
                stylesController = new StylesController(Resources, corelApp);

            }
            catch
            {
                MessageBox.Show("VGCore Erro");
            }
            //подписка на события комбобоксов и добавление значений
            Velosity.Loaded += CloseCmbBx;
            Overcut.Loaded += CloseCmbBx;
            Velosity.DropDownClosed += CloseCmbBx;
            Overcut.DropDownClosed += CloseCmbBx;
            for (int vel = 100; vel < 1010; vel += 10)
                    {
                        Velosity.Items.Add(new ComboItems
                        {
                            Value = vel,
                            pref = vel + " мм/сек"
                        });
                    }
            Velosity.Items.Add(new ComboItems
            {
                Value = -1,
                pref = "Использовать значение устройства"
            });

            for (int ovr = 1; ovr < 11; ovr++)
            {
                Overcut.Items.Add(new ComboItems
                {
                    Value = ovr,
                    pref = ovr / 10m + "мм"
                });
            }
            Overcut.Items.Add(new ComboItems
            {
                Value = -1,
                pref = "Использовать значение устройства"
            });

           

            //подставляем значения переменных из файла
            Velosity.SelectedItem = Velosity.Items.Cast<ComboItems>().First(x => x.Value == settings.velosity);
            Overcut.SelectedItem = Overcut.Items.Cast<ComboItems>().First(x => x.Value == settings.overcut);
            Smothing.IsChecked = settings.smothing;
            Barcode2.IsChecked = settings.barc2;
            Path.Text = settings.path_plt;
            PlxCheck.IsChecked = settings.plx;
            Color.Text = string.Join(",", settings.color_name);
            NameDoc.IsChecked = settings.doc_name;
            UpBleed.Text = settings.up_bleed.ToString() + " мм";
            DownBleed.Text = settings.down_bleed.ToString() + " мм";
            LeftBleed.Text = settings.left_bleed.ToString() + " мм";
            RightBleed.Text = settings.right_bleed.ToString() + " мм";
            //подписка на события текстбокса
            UpBleed.GotFocus += BleedChanged;
            UpBleed.LostFocus += BleedChanged;
            UpBleed.PreviewTextInput += Bleed_TextChanged;
            DownBleed.GotFocus += BleedChanged;
            DownBleed.LostFocus += BleedChanged;
            DownBleed.PreviewTextInput += Bleed_TextChanged;
            LeftBleed.GotFocus += BleedChanged;
            LeftBleed.LostFocus += BleedChanged;
            LeftBleed.PreviewTextInput += Bleed_TextChanged;
            RightBleed.GotFocus += BleedChanged;
            RightBleed.LostFocus += BleedChanged;
            RightBleed.PreviewTextInput += Bleed_TextChanged;
        }

       

        private void CloseCmbBx(object sender,EventArgs e)
        {
            ComboBox cb = (ComboBox)sender; 
            double txt = DefineWidthText(cb);
            cb.Width = Math.Max(cb.MinWidth, txt + 30);
            double maxComboWidth = Math.Max(Velosity.Width, Overcut.Width);
            double requiredWidth = 120 + maxComboWidth + 30;
            this.Width = Math.Max(this.MinWidth, requiredWidth);
        }

        public double DefineWidthText(ComboBox cb)
        {
            var text = new FormattedText(cb.SelectedItem.ToString(),
            System.Globalization.CultureInfo.CurrentCulture,
            System.Windows.FlowDirection.LeftToRight,
            new Typeface(cb.FontFamily, cb.FontStyle,
                         cb.FontWeight, cb.FontStretch),
            cb.FontSize,
            Brushes.Black,
            VisualTreeHelper.GetDpi(cb).PixelsPerDip);
            return text.Width;
        }

        private void Bleed_TextChanged(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !e.Text.All(char.IsDigit);
        }

        
        private void BleedChanged(object sender, RoutedEventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if (tb.IsFocused)
            {
                // Убираем единицы измерения, но сохраняем число.
                tb.Text = tb.Text.Replace(" мм", "").Trim();

                // Выделяем всё значение, чтобы его можно было
                // сразу заменить вводом нового числа.
                tb.SelectAll();
            }
            else
            {
                if (string.IsNullOrWhiteSpace(tb.Text))
                {
                    // Если поле действительно пустое, возвращаем
                    // соответствующее значение из настроек.
                    if (tb == UpBleed)
                        tb.Text = settings.up_bleed.ToString() + " мм";
                    else if (tb == DownBleed)
                        tb.Text = settings.down_bleed.ToString() + " мм";
                    else if (tb == LeftBleed)
                        tb.Text = settings.left_bleed.ToString() + " мм";
                    else if (tb == RightBleed)
                        tb.Text = settings.right_bleed.ToString() + " мм";
                }
                else
                {
                    // Сохраняем введённое пользователем число.
                    tb.Text = tb.Text.Replace(" мм", "").Trim() + " мм";
                }
            }
          

        }
        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();

        }

       
        public void Ok_click(object sender, RoutedEventArgs e)
        {
            try
            {
                //обновляем переменнные
                settings.velosity = ((ComboItems)Velosity.SelectedItem).Value; ;
                settings.overcut = (int)((decimal)((ComboItems)Overcut.SelectedItem).Value);
                settings.smothing = Smothing.IsChecked == true;
                settings.barc2 = Barcode2.IsChecked == true;
                settings.path_plt = Path.Text;
                settings.color_name = Color.Text.Split(new char[] { ',' });
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
                    xsz.Serialize(fs, settings);
                }
                MessageBox.Show("Настройки резки сохранены!", "Добавление меток и экспорт plt", 
                                 MessageBoxButton.OK,
                                 MessageBoxImage.Information);
                Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
