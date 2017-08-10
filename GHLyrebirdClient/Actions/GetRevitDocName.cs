using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lyrebird
{
    public static class Actions
    {
        public static string GetRevitDocName(Autodesk.Revit.UI.UIApplication uiApp)
        {
            try
            {
                return uiApp.ActiveUIDocument.Document.Title;
            }
            catch
            {
                return null;
            }
        }
    }
}
