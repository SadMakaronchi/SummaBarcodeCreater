using System.Threading;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Progress_bar
    {
        private Thread uiThread;
        private Progress win;
        private Application corelApp;

        
        public void Progress(Application app) //запуск полосы прогресса экспорта в отдельнном потоке
        {
            this.corelApp = app;
            uiThread = new Thread(() =>
            {
                win = new Progress();
                win.Show();
                win.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                System.Windows.Threading.Dispatcher.Run();
            });
            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.Start();
        }


        
        public void Progress_stop()//остановка полосы прогресса экспорта
        {
            if (win != null)
            {
                win.Dispatcher.Invoke(() =>
                {
                    win.Close();
                    win = null;
                });
            }
        }
    }
}
