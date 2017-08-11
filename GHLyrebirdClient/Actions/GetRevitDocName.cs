using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lyrebird
{
    public class GetRevitDocName
    {
        public static bool Command(Autodesk.Revit.UI.UIApplication uiApp, Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            try
            {
                string docName = uiApp.ActiveUIDocument.Document.Title;
                outputs = new Dictionary<string, object>{{"docName", docName}};
                return true;
            }
            catch
            {
                outputs = null;
                return false;
            }
        }

        public static Guid CommandGuid => new Guid("a7638694-0ad6-4e40-8e2b-a26a427c0678");
    }
}