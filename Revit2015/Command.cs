using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace LMNA.Lyrebird
{
    [Transaction(TransactionMode.Manual)]
    class ServerToggle : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                RevitServerApp.Instance.Toggle();
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
