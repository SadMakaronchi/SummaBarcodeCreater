using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace SummaBarcodeCreater
{
    public class Settings_define
    {
        public string folder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel");
        public string path_settings = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SummaPanel", "SettingsCutSumma.xml");

        
        public Settings Get_settings()
        {
            // получаем настройки из файла
            Settings settings = new Settings();
            if (Directory.Exists(folder) == false)
            {
                try
                {
                    Directory.CreateDirectory(folder);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Нет прав на создание папки для хранения файла с настройками!Пожалуйста запустите CorelDraw с правами администратора");
                    return settings;
                }
                try
                {
                    using (FileStream fs = File.Create(path_settings))
                    {
                        XmlSerializer xsz = new XmlSerializer(typeof(Settings));
                        xsz.Serialize(fs, settings);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка создания файла настроек!" + ex);
                    return settings;
                }
            }
            else
            {
                using (FileStream fs = File.OpenRead(path_settings))
                {
                    XmlSerializer xsz = new XmlSerializer(typeof(Settings));
                    settings = (Settings)xsz.Deserialize(fs);
                }
            }

            return settings;


        }

    }
}
