using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using corel = Corel.Interop.VGCore;

namespace SummaBarcodePlugin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SummaPlugin
    {
        private corel.Application app;

        public SummaPlugin(object application)
        {
            app = (corel.Application)application;

            // 1. Регистрируем команду
            AddCommand();

            // 2. Создаём тулбар и кнопку
            AddToolbarButton();
        }

        private void AddCommand()
        {
            try
            {
                // ID команды = имя метода
                app.AddPluginCommand(
                    "SummaBarcodeCreate",       // внутренний ID
                    "Создать штрихкод",         // название в UI
                    "Создаёт штрихкод Summa",   // tooltip
                    
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка AddPluginCommand: " + ex.Message);
            }
        }

        private void AddToolbarButton()
        {
            try
            {
                corel.CommandBar bar;

                // Проверяем, есть ли тулбар
                try
                {
                    bar = app.CommandBars["SummaTools"];
                }
                catch
                {
                    bar = app.CommandBars.Add("SummaTools", corel.CdrBarPosition.cdrBarTop, false);
                }

                // Проверяем, есть ли кнопка
                bool exists = false;
                foreach (corel.CommandBarControl ctrl in bar.Controls)
                {
                    if (ctrl.ID == "SummaBarcodeCreate")
                    {
                        exists = true;
                        break;
                    }
                }

                if (!exists)
                {
                    var btn = bar.Controls.Add(corel.CdrControlType.cdrControlButton,
                                               "SummaBarcodeCreate", "", true);
                    btn.Caption = "Summa Barcode";
                    btn.TooltipText = "Создать штрихкод Summa";
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка AddToolbarButton: " + ex.Message);
            }
        }

        // Этот метод вызывается при нажатии кнопки
        public void SummaBarcodeCreate()
        {
            System.Windows.Forms.MessageBox.Show("Кнопка нажата! Ваш код здесь.");
        }
    }
}
