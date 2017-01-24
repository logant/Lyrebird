using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    public class LBHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            try
            {
                // Do something
            }
            catch (Exception ex)
            {
                TaskDialog dlg = new TaskDialog("Lyrebird Error");
                dlg.TitleAutoPrefix = false;
                dlg.MainInstruction = "Error";
                dlg.MainInstruction = ex.Message;
                dlg.Show();
            }
        }

        public string GetName()
        {
            return "Lyrebird Server";
        }
    }
}
