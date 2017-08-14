using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Lyrebird
{
    public class GetRevitDocName
    {
        public GetRevitDocName(string installPath)
        {
            try
            {
                Assembly.LoadFrom(installPath + "\\RevitAPI.dll");
                Assembly.LoadFrom(installPath + "\\RevitAPIUI.dll");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error:\n" + e.Message);
            }
        }
        
        public bool Command(Autodesk.Revit.UI.UIApplication uiApp, Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            try
            {
                string docName = uiApp.ActiveUIDocument.Document.Title;
                outputs = new Dictionary<string, object>{{"docName", docName}};
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error:\n" + ex.Message);
                outputs = null;
                return false;
            }
        }

        public static Guid CommandGuid => new Guid("8a9ce1a6-9861-49f0-8b3b-535779d197b5");
    }
}