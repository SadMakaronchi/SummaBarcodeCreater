using System;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Entry
    {
        private Application corelApp;
        
        public void Begin(Application app) //запускаем основной код плагина
        {

            corelApp = app;
            if (corelApp == null)
                throw new InvalidOperationException("Corel не инициализирован");
            try
            {
                var macros = new Macros(); //непосредственно макрос
                macros.StartMacros(corelApp);
            }
            finally
            {
                var progress = new Progress_bar();
                progress.Progress_stop();
            }


        }

    }

}
