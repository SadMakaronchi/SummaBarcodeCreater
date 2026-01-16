using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using corel = Corel.Interop.VGCore;

namespace SettingCutSumma
{
    [ComVisible(true)]
    [Guid("b3f7d9a3-02f6-4259-9b13-28c2c8070cfa")]
    public class Entry
    {
        public corel.Application crl = new corel.Application();
        public void Start()
        {

            var win = new MainWindow(crl);
            win.Show();
        }
        private Thread uiThread;
        private progress win;
        public void Progress()
        {
            if (uiThread != null && uiThread.IsAlive)
            {
                win.Dispatcher.Invoke(() =>
                {
                    if (win.WindowState == WindowState.Minimized)
                        win.WindowState = WindowState.Normal;
                        win.Activate();
                });
                return;
            }
            uiThread = new Thread(() =>
            {
                win = new progress();
                win.Closed += (s, e) => win.Dispatcher.InvokeShutdown();
                win.Show();
                System.Windows.Threading.Dispatcher.Run();
            });
            
            uiThread.SetApartmentState(ApartmentState.STA); // обязательно для WPF
            uiThread.IsBackground = true;
            uiThread.Start();
        }

        

        public void Progress_stop()
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
