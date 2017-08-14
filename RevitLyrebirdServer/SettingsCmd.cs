using System;
using System.Diagnostics;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    [Transaction(TransactionMode.Manual)]
    public class SettingsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Build the settings form.
                Process proc = Process.GetCurrentProcess();
                IntPtr handle = proc.MainWindowHandle;
                LyrebirdSettings window = new LyrebirdSettings();
                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = handle;
                window.ShowDialog();


                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
