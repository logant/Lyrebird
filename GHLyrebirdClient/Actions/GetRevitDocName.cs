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
//                System.Windows.Forms.MessageBox.Show("In the GetDocName constructor.");
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
//                System.Windows.Forms.MessageBox.Show("In the GetDocName constructor.");
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

        public static Guid CommandGuid => new Guid("a7638694-0ad6-4e40-8e2b-a26a427c0678");
    }
}