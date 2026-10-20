using System;
using System.Windows;
using SummaBarcodeCreater;
using Application = Corel.Interop.VGCore.Application;

namespace CGS
{
    //этот класс видит корел при загрузке, важно заменить расширение dll на CorelAddon и испоьзовать пространство имён CGS
    public class Addon
    {
        public static Application corelApp;
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
                dataSourceFactory = new DataSourceFactory(corelApp);
                dataSourceFactory.AddDataSource(
                DataSourceName,typeof(SummaBarcodeCreaterDataSource));
                dataSourceFactory.Register(); //регистрация фабрики данных для связывания интерфейса с методами плагина
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка регистрации DataSource" + ex);
            }
        }
    }
}