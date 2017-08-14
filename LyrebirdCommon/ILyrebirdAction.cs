using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace Lyrebird
{
    public abstract class ILyrebirdAction
    {
        public abstract bool Command(Autodesk.Revit.UI.UIApplication uiApp, Dictionary<string, object> inputs,
            out Dictionary<string, object> outputs);

        public ILyrebirdAction(string installPath)
        {
            
        }
    }
}
