using System.IO;
using Application = Corel.Interop.VGCore.Application;
using Corel.Interop.VGCore;
namespace SummaBarcodeCreater
{
    public class ExportToPDF
    {
        public Application corelApp;
        public Numbers num;
        public Settings set;
        public void BeginExport(Application app, Numbers numbers,Settings settings)
        {
            corelApp = app;
            num = numbers;
            string path = settings.path_plt;
            string name = Path.Combine(path, numbers.down_number.ToString() + ".pdf");
            var expt = corelApp.ActiveDocument.ExportEx(name, cdrFilter.cdrPDF, cdrExportRange.cdrSelection);
        }
    }
}
