using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    [ServiceBehavior]
    public class LBService : ILyrebirdService
    {
        private static readonly object Locker = new object();
        private const int WAIT_TIMEOUT = 1000;

        public bool LbAction(Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            bool allowMessages = false;
            outputs = null;
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    if (string.IsNullOrEmpty(Properties.Settings.Default.LBComponentPath) ||
                        !Directory.Exists(Properties.Settings.Default.LBComponentPath))
                    {
                        TaskDialog dlg = new TaskDialog("Error");
                        dlg.TitleAutoPrefix = false;
                        dlg.MainInstruction = "Lyrebird Components Folder Not Set";
                        dlg.MainContent =
                            "An attempt was made to use Lyrebird before the Lyrebird Components folder was set. " + 
                            "Please use the Lyrebird settings to specify where the Lyrebird components are. Hint, " + 
                            "these are typically in the Grasshopper Libraries folder at %APPDATA%\\Grasshopper\\Libraries\\Lyrebird";
                        dlg.Show();
                        return false;
                    }
                    
                    LBApp.Handler.Inputs = inputs;
                    LBApp.ExEvent.Raise();

                    outputs = LBApp.Handler.Outputs;
                    if (outputs == null)
                        return false;

                    return true;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("I threw an exception.  :\n" + ex.Message);
                    //do nothing
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
            }

            return (outputs == null);
        }

        public bool Ping()
        {
            return true;
        }
    }
}