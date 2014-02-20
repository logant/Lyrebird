using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage;
using System.Diagnostics;

using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird
{
    [Transaction(TransactionMode.Manual)]
    class RemoveDataCmd : IExternalCommand
    {
        private readonly Guid instanceSchemaGUID = new Guid("9ab787e0-1660-40b7-9453-94e1043b58db");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Autodesk.Revit.UI.Selection.SelElementSet elemSet = commandData.Application.ActiveUIDocument.Selection.Elements;
                int itemCount = 0;
                if (elemSet.Size > 0)
                {
                    Schema schema = Schema.Lookup(instanceSchemaGUID);
                    if (schema == null)
                    {
                        TaskDialog.Show("Error", "Could not find any Lyrebird Data");
                    }
                    else
                    {
                        using (Transaction trans = new Transaction(commandData.Application.ActiveUIDocument.Document, "Lyrebird - Remove Data"))
                        {
                            trans.Start();

                            foreach (Element e in elemSet)
                            {
                                try
                                {
                                    e.DeleteEntity(schema);
                                    itemCount++;
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }
                            trans.Commit();
                            TaskDialog.Show("Message", "Successfully removed Lyrebird data from " + itemCount.ToString() + " elements.");
                        }
                    }
                }
                else
                {
                    TaskDialog.Show("Error", "No elements selected.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

        }
    }
}
