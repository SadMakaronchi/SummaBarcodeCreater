using Corel.Interop.VGCore;
using System.Runtime.InteropServices;
using System.Windows;
using Application = Corel.Interop.VGCore.Application;
namespace SummaBarcodeCreater
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("f9ae43c4-9634-41ac-be3b-9ed3e30bc349")]
    public class SummaBarcodeCreaterDataSource : BaseDataSource
    {
        public SummaBarcodeCreaterDataSource(DataSourceProxy proxy,
                               Application corelApp)
                             : base(proxy, corelApp)
        {
        }

        public void RunMacros()
        {
            var entry = new Entry();
            entry.Begin(CorelApp);
        }

        public void Settings()
        {
            var set = new Settings_win();
            set.Start(CorelApp);
        }
    }
}
