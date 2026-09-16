using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Threading;
using Application = Corel.Interop.VGCore.Application;
namespace SummaBarcodeCreater
{
    public class Settings_win
    {
        private Thread uiThread;
        private MainWindow win;

        public void Start(Application app) //запуск окна настроек в отдельном потоке
        {
            Settings_define settings_define = new Settings_define();
            Settings settings = settings_define.Get_settings();
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = settings.path_plt,
                Title = "Выберите папку для сохранения файла резки в формате plt"
            };

            uiThread = new Thread(() =>
            {
                win = new MainWindow(app,dialog);
                app.OnApplicationEvent += CorelApp_OnApplicationEvent;
                win.Show();
                win.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                System.Windows.Threading.Dispatcher.Run();
            });
            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.Start();
        }
        private void CorelApp_OnApplicationEvent(
        string eventName,
        ref object[] parameters)
        {
            if (eventName != "WorkspaceChanged" &&
                eventName != "OnColorSchemeChanged")
                return;

            if (win == null)
                return;

            win.Dispatcher.BeginInvoke(
                new Action(() =>
                {
                    win.stylesController?.LoadThemeFromPreference();
                }));
        }
    }
}

