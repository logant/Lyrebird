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
    class SelectRunElementsCmd : IExternalCommand
    {
        private readonly Guid instanceSchemaGUID = new Guid("9ab787e0-1660-40b7-9453-94e1043b58db");
        Document doc;
        List<RunCollection> allRuns = new List<RunCollection>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Find all of the runs and organize them
                doc = commandData.Application.ActiveUIDocument.Document;
                allRuns = FindExisting();
                
                if (allRuns != null && allRuns.Count > 0)
                {
                    RevitServerApp._app.ShowSelectionForm(allRuns, commandData.Application.ActiveUIDocument);
                }
                else
                {
                    TaskDialog.Show("Message", "No existing run elements found.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error", ex.Message);
                return Result.Failed;
            }
        }

        private List<RunCollection> FindExisting()
        {
            // Find existing elements with a matching GUID from the GH component.
            List<RunCollection> collectedRuns = new List<RunCollection>();

            Schema instanceSchema = Schema.Lookup(instanceSchemaGUID);
            if (instanceSchema == null)
            {
                return collectedRuns;
            }

            ElementIsElementTypeFilter filter = new ElementIsElementTypeFilter(false);

            FilteredElementCollector fiCollector = new FilteredElementCollector(doc);
            fiCollector.OfClass(typeof(FamilyInstance));
            fiCollector.ToElements();

            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfCategory(BuiltInCategory.OST_Walls);
            wallCollector.OfClass(typeof(Wall));
            wallCollector.ToElements();

            FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
            floorCollector.OfCategory(BuiltInCategory.OST_Floors);
            floorCollector.OfClass(typeof(Floor));
            floorCollector.ToElements();

            FilteredElementCollector roofCollector = new FilteredElementCollector(doc);
            roofCollector.OfCategory(BuiltInCategory.OST_Roofs);
            roofCollector.OfClass(typeof(RoofBase));
            roofCollector.ToElements();

            List<Element> elemCollector = new List<Element>();
            foreach (Element e in fiCollector)
            {
                elemCollector.Add(e);
            }
            foreach (Element e in wallCollector)
            {
                elemCollector.Add(e);
            }
            foreach (Element e in floorCollector)
            {
                elemCollector.Add(e);
            }
            foreach (Element e in roofCollector)
            {
                elemCollector.Add(e);
            }

            //FilteredElementCollector elemCollector = new FilteredElementCollector(doc);
            //elemCollector.WherePasses(filter).ToElements();
            
            // First, find all of the unique componentGUID's that are in the Revit file.
            List<string> instanceIds = new List<string>();
            foreach (Element e in elemCollector)
            {
                if (e.Category.Id.IntegerValue == -2000011)
                {
                    try
                    {
                        Wall w = e as Wall;
                        if (w != null)
                        {
                            Entity entity = w.GetEntity(instanceSchema);
                            if (entity.IsValid())
                            {
                                Field f = instanceSchema.GetField("InstanceID");
                                string tempId = entity.Get<string>(f);
                                if (!instanceIds.Contains(tempId))
                                {
                                    instanceIds.Add(tempId);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                }
                else if (e.Category.Id.IntegerValue == -2000032)
                {
                    try
                    {
                        Floor flr = e as Floor;
                        if (flr != null)
                        {
                            Entity entity = flr.GetEntity(instanceSchema);
                            if (entity.IsValid())
                            {
                                Field f = instanceSchema.GetField("InstanceID");
                                string tempId = entity.Get<string>(f);
                                if (!instanceIds.Contains(tempId))
                                {
                                    instanceIds.Add(tempId);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                }
                else if (e.Category.Id.IntegerValue == -2000035)
                {
                    try
                    {
                        RoofBase r = e as RoofBase;
                        if (r != null)
                        {
                            Entity entity = r.GetEntity(instanceSchema);
                            if (entity.IsValid())
                            {
                                Field f = instanceSchema.GetField("InstanceID");
                                string tempId = entity.Get<string>(f);
                                if (!instanceIds.Contains(tempId))
                                {
                                    instanceIds.Add(tempId);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        FamilyInstance fi = e as FamilyInstance;
                        if (fi != null)
                        {
                            Entity entity = fi.GetEntity(instanceSchema);
                            if (entity.IsValid())
                            {
                                Field f = instanceSchema.GetField("InstanceID");
                                string tempId = entity.Get<string>(f);
                                if (!instanceIds.Contains(tempId))
                                {
                                    instanceIds.Add(tempId);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                }
            }
            
            // Create a runCollection for each guid
            foreach (string id in instanceIds)
            {
                RunCollection rc = new RunCollection();
                List<Runs> tempRuns = new List<Runs>();
                
                // Find the number of runs per instanceId
                List<int> runIds = new List<int>();
                foreach (Element e in elemCollector)
                {
                    if (e.Category.Id.IntegerValue == -2000011)
                    {
                        // Walls
                        try
                        {
                            Wall w = e as Wall;
                            if (w != null)
                            {
                                Entity entity = w.GetEntity(instanceSchema);
                                if (entity.IsValid())
                                {
                                    Field f = instanceSchema.GetField("InstanceID");
                                    string tempId = entity.Get<string>(f);
                                    if (tempId == id)
                                    {
                                        rc.ComponentGuid = new Guid(tempId);
                                        f = instanceSchema.GetField("RunID");
                                        int tempRunId = entity.Get<int>(f);

                                        if (!runIds.Contains(tempRunId))
                                        {
                                            runIds.Add(tempRunId);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error", ex.Message);
                        }
                    }
                    else if (e.Category.Id.IntegerValue == -2000032)
                    {
                        // Floors
                        try
                        {
                            Floor flr = e as Floor;
                            if (flr != null)
                            {
                                Entity entity = flr.GetEntity(instanceSchema);
                                if (entity.IsValid())
                                {
                                    Field f = instanceSchema.GetField("InstanceID");
                                    string tempId = entity.Get<string>(f);
                                    if (tempId == id)
                                    {
                                        rc.ComponentGuid = new Guid(tempId);
                                        f = instanceSchema.GetField("RunID");
                                        int tempRunId = entity.Get<int>(f);

                                        if (!runIds.Contains(tempRunId))
                                        {
                                            runIds.Add(tempRunId);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error", ex.Message);
                        }
                    }
                    else if (e.Category.Id.IntegerValue == -2000035)
                    {
                        // Roofs
                        try
                        {
                            RoofBase r = e as RoofBase;
                            if (r != null)
                            {
                                Entity entity = r.GetEntity(instanceSchema);
                                if (entity.IsValid())
                                {
                                    Field f = instanceSchema.GetField("InstanceID");
                                    string tempId = entity.Get<string>(f);
                                    if (tempId == id)
                                    {
                                        rc.ComponentGuid = new Guid(tempId);
                                        f = instanceSchema.GetField("RunID");
                                        int tempRunId = entity.Get<int>(f);

                                        if (!runIds.Contains(tempRunId))
                                        {
                                            runIds.Add(tempRunId);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error", ex.Message);
                        }
                    }
                    else
                    {
                        // Other non-system families
                        try
                        {
                            FamilyInstance fi = e as FamilyInstance;
                            if (fi != null)
                            {
                                Entity entity = fi.GetEntity(instanceSchema);
                                if (entity.IsValid())
                                {
                                    Field f = instanceSchema.GetField("InstanceID");
                                    string tempId = entity.Get<string>(f);
                                    if (tempId == id)
                                    {
                                        rc.ComponentGuid = new Guid(tempId);
                                        f = instanceSchema.GetField("RunID");
                                        int tempRunId = entity.Get<int>(f);

                                        if (!runIds.Contains(tempRunId))
                                        {
                                            runIds.Add(tempRunId);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error", ex.Message);
                        }
                    }
                }
                
                foreach (int i in runIds)
                {
                    List<int> runElemIds = new List<int>();
                    Runs run = new Runs();
                    foreach (Element e in elemCollector)
                    {
                        try
                        {
                            if (e.Category.Id.IntegerValue == -2000011)
                            {
                                // Walls
                                try
                                {
                                    Wall w = e as Wall;
                                    if (w != null)
                                    {
                                        Entity entity = w.GetEntity(instanceSchema);
                                        if (entity.IsValid())
                                        {
                                            Field f = instanceSchema.GetField("InstanceID");
                                            string tempId = entity.Get<string>(f);
                                            if (tempId == id)
                                            {
                                                f = instanceSchema.GetField("RunID");
                                                int tempRunId = entity.Get<int>(f);

                                                if (tempRunId == i)
                                                {
                                                    if (run.RunName == null || run.RunName == string.Empty)
                                                    {
                                                        run.RunId = tempRunId;
                                                        run.RunName = "Run" + tempRunId.ToString();
                                                        run.FamilyType = w.Category.Name + " : " + w.WallType.Name;
                                                    }
                                                    runElemIds.Add(w.Id.IntegerValue);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }
                            else if (e.Category.Id.IntegerValue == -2000032)
                            {
                                // Floors
                                try
                                {
                                    Floor flr = e as Floor;
                                    if (flr != null)
                                    {
                                        Entity entity = flr.GetEntity(instanceSchema);
                                        if (entity.IsValid())
                                        {
                                            Field f = instanceSchema.GetField("InstanceID");
                                            string tempId = entity.Get<string>(f);
                                            if (tempId == id)
                                            {
                                                f = instanceSchema.GetField("RunID");
                                                int tempRunId = entity.Get<int>(f);

                                                if (tempRunId == i)
                                                {
                                                    if (run.RunName == null || run.RunName == string.Empty)
                                                    {
                                                        run.RunId = tempRunId;
                                                        run.RunName = "Run" + tempRunId.ToString();
                                                        run.FamilyType = flr.Category.Name + " : " + flr.FloorType.Name;
                                                    }
                                                    runElemIds.Add(flr.Id.IntegerValue);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }
                            else if (e.Category.Id.IntegerValue == -2000035)
                            {
                                // Roofs
                                try
                                {
                                    RoofBase r = e as RoofBase;
                                    if (r != null)
                                    {
                                        Entity entity = r.GetEntity(instanceSchema);
                                        if (entity.IsValid())
                                        {
                                            Field f = instanceSchema.GetField("InstanceID");
                                            string tempId = entity.Get<string>(f);
                                            if (tempId == id)
                                            {
                                                f = instanceSchema.GetField("RunID");
                                                int tempRunId = entity.Get<int>(f);

                                                if (tempRunId == i)
                                                {
                                                    if (run.RunName == null || run.RunName == string.Empty)
                                                    {
                                                        run.RunId = tempRunId;
                                                        run.RunName = "Run" + tempRunId.ToString();
                                                        run.FamilyType = r.Category.Name + " : " + r.RoofType.Name;
                                                    }
                                                    runElemIds.Add(r.Id.IntegerValue);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }
                            else
                            {
                                // Other non-system families
                                try
                                {
                                    FamilyInstance fi = e as FamilyInstance;
                                    if (fi != null)
                                    {
                                        Entity entity = fi.GetEntity(instanceSchema);
                                        if (entity.IsValid())
                                        {
                                            Field f = instanceSchema.GetField("InstanceID");
                                            string tempId = entity.Get<string>(f);
                                            if (tempId == id)
                                            {
                                                f = instanceSchema.GetField("RunID");
                                                int tempRunId = entity.Get<int>(f);

                                                if (tempRunId == i)
                                                {
                                                    if (run.RunName == null || run.RunName == string.Empty)
                                                    {
                                                        run.RunId = tempRunId;
                                                        run.RunName = "Run" + tempRunId.ToString();
                                                        run.FamilyType = fi.Symbol.Family.Name + " : " + fi.Symbol.Name;
                                                    }
                                                    runElemIds.Add(fi.Id.IntegerValue);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error", ex.Message);
                        }
                    }
                    run.ElementIds = runElemIds;
                    tempRuns.Add(run);
                }
                rc.Runs = tempRuns;
                collectedRuns.Add(rc);
            }
            return collectedRuns;
        }
    }
}
