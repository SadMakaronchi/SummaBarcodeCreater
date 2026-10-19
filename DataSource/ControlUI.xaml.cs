using System;
using System.Windows.Controls;
using Corel.Interop.VGCore;
namespace SummaBarcodeCreater
{
    
    public partial class ControlUI : UserControl
    {
        public static Application corelApp = null;
        public static IntPtr corelHandle;
        public const string DataSourceName = "SummaBarcodeCreaterDS";
        
        public ControlUI(object app)
        {

            try
            {
                corelApp = app as Application;
                corelHandle = new IntPtr(corelApp.AppWindow.Handle);
                var dsf = new DataSourceFactory(corelApp);
                dsf.AddDataSource(DataSourceName, typeof(SummaBarcodeCreaterDataSource));
                dsf.Register();
            }
            catch (Exception)
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }
        }
    }
}
