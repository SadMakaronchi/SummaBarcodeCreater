using System;
using System.Runtime.InteropServices;
using System.Threading;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Progress_bar
    {
        private Thread uiThread;
        public Progress win;
        private Application corelApp;
        private IntPtr hwnd;

        private ManualResetEventSlim windowReady =
            new ManualResetEventSlim(false);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(
            IntPtr hWnd,
            string text,
            string caption,
            uint type);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public void Progress(Application app)
        {
            corelApp = app;

            hwnd = GetForegroundWindow();

            // На случай повторного запуска
            windowReady.Reset();

            uiThread = new Thread(() =>
            {
                win = new Progress(app);

                win.Show();

                win.Dispatcher.Invoke(
                    () => { },
                    System.Windows.Threading.DispatcherPriority.Render);

                // Сообщаем основному потоку,
                // что окно уже создано и показано
                windowReady.Set();

                System.Windows.Threading.Dispatcher.Run();
            });

            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.Start();

            // Ждём создания окна
            windowReady.Wait();
        }

        public void ShowMessage(
            string text,
            string caption = "Добавление меток и экспорт plt")
        {
            // Сначала закрываем Progress
            Progress_stop();

            // Потом показываем MessageBox
            MessageBox(
                hwnd,
                text,
                caption,
                0x30);
        }

        public void Progress_stop()
        {
            if (win != null)
            {
                win.Dispatcher.Invoke(() =>
                {
                    if (win != null)
                    {
                        win.Close();
                        win = null;
                    }
                });
            }
        }
    }
}