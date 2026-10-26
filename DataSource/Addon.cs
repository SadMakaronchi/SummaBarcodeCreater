using Corel.Interop.VGCore;
using SummaBarcodeCreater;
using System;
using System.ComponentModel.Design;
using System.Windows;
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
                //dataSourceFactory = new DataSourceFactory(corelApp);
                //dataSourceFactory.AddDataSource(
                //DataSourceName, typeof(SummaBarcodeCreaterDataSource));
                //dataSourceFactory.Register(); //регистрация фабрики данных для связывания интерфейса с методами плагина

                corelApp.AddPluginCommand("AddBarcode");
                var cmb = corelApp.CommandBars.Add("SummaBarcodeCreater", Corel.Interop.VGCore.cuiBarPosition.cuiBarFloating);
                var bt1 = cmb.Controls.AddCustomButton("ab489730-8791-45d2-a825-b78bbe0d6a5d", "AddBarcode");
                bt1.Caption = "AddBarcode";

                corelApp.OnPluginCommand += CorelApp_OnPluginCommand;
                corelApp.OnUpdatePluginCommand += CorelApp_OnUpdatePluginCommand;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка регистрации DataSource" + ex);
            }


        }

        private void CorelApp_OnPluginCommand(string commandID)
        {
            if (commandID == "AddBarcode")
            {
                var entry = new Entry();
                entry.Begin(corelApp);
            }
        }

        private void CorelApp_OnUpdatePluginCommand(string commandID,
                                                    ref bool enabled,
                                                    ref cdrCommandCheckState checkedState)
        {
            if (commandID == "AddBarcode")
                enabled = true;
        }
    }
}