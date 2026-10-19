using System;
using System.Windows;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Addon
    {
        public static Application corelApp;
        public static IntPtr corelHandle;

        public const string DataSourceName = "SummaBarcodeCreaterDS";

        private readonly DataSourceFactory dataSourceFactory;

        public Addon(Application app)
        {
            try
            {
                corelApp = app;

                if (corelApp == null)
                {
                    throw new InvalidOperationException("Corel Application равен null");
                }
                    corelHandle = new IntPtr(corelApp.AppWindow.Handle);
                    dataSourceFactory = new DataSourceFactory(corelApp);
                    dataSourceFactory.AddDataSource(
                    DataSourceName, typeof(SummaBarcodeCreaterDataSource));
                    dataSourceFactory.Register();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка регистрации DataSource" + ex);
            }
        }
    }
}