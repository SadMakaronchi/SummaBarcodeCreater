using Corel.Interop.VGCore;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace SummaBarcodeCreater
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class BaseDataSource: INotifyPropertyChanged
    {
        protected DataSourceProxy m_AppProxy;
        protected Application CorelApp;
        public BaseDataSource(DataSourceProxy proxy, Application corelApp)
        {
            m_AppProxy = proxy;
            CorelApp = corelApp;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            try
            {
                this.m_AppProxy.UpdateListeners(propertyName);
                if (PropertyChanged != null)
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
            catch { }
        }
    }
}
