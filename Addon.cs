using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Corel.Interop.VGCore.Application;

namespace SummaBarcodeCreater
{
    public class Addon
    {
        public Addon(Application app) 
        {
            var cmbr =  app.CommandBars.Add("testMenu");
            var btn = cmbr.Controls.AddCustomButton("e69f024e-30d4-4ed0-86c5-f8615e307eb0", "button");
        }
    }
}
