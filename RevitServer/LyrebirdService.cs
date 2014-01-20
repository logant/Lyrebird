using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using LMNA.Lyrebird.LyrebirdCommon;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace LMNA.Lyrebird
{
    [ServiceBehavior]
    public class LyrebirdService : ILyrebirdService
    {
        private string currentDocName = "NULL";
        private List<RevitObject> familyNames = new List<RevitObject>();
        private List<string> typeNames = new List<string>();
        private List<RevitParameter> parameters = new List<RevitParameter>();
        private List<LyrebirdId> uniqueIds = new List<LyrebirdId>();

        private Guid instanceSchemaGUID = new Guid("9ab787e0-1660-40b7-9453-94e1043b58db");

        private static readonly object _locker = new object();

        private const int WAIT_TIMEOUT = 200;

        public List<RevitObject> GetFamilyNames()
        {
            familyNames.Add(new RevitObject("NULL", "NULL"));
            lock (_locker)
            {
                try
                {
                    UIApplication uiApp = LMNA.Lyrebird.RevitServerApp.UIApp;
                    familyNames = new List<RevitObject>();

                    // Get all standard wall families
                    FilteredElementCollector familyCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    familyCollector.OfClass(typeof(Family));
                    List<RevitObject> families = new List<RevitObject>();
                    foreach (Family f in familyCollector)
                    {
                        RevitObject ro = new RevitObject(f.FamilyCategory.Name, f.Name);
                        families.Add(ro);
                    }

                    // Add System families
                    RevitObject wallObj = new RevitObject("Walls", "Basic Wall");
                    families.Add(wallObj);
                    RevitObject curtainObj = new RevitObject("Walls", "Curtain Wall");
                    families.Add(curtainObj);
                    RevitObject stackedObj = new RevitObject("Walls", "Stacked Wall");
                    families.Add(stackedObj);
                    RevitObject floorObj = new RevitObject("Floors", "Floor");
                    families.Add(floorObj);
                    RevitObject roofObj = new RevitObject("Roofs", "Roof");
                    families.Add(roofObj);

                    families.Sort((x, y) => String.Compare(x.FamilyName, y.FamilyName));
                    familyNames = families;
                }
                catch
                {

                }
            Monitor.Wait(_locker, WAIT_TIMEOUT);
            }
            return familyNames;
        }

        public List<string> GetTypeNames(RevitObject revitFamily)
        {
            typeNames.Add("NULL");
            lock (_locker)
            {
                try
                {
                    UIApplication uiApp = LMNA.Lyrebird.RevitServerApp.UIApp;
                    var doc = uiApp.ActiveUIDocument.Document;
                    typeNames = new List<string>();
                    List<string> types = new List<string>();
                    if (revitFamily.Category == "Walls")
                    {
                        // get wall types
                        FilteredElementCollector wallCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        wallCollector.OfClass(typeof(WallType));
                        wallCollector.OfCategory(BuiltInCategory.OST_Walls);
                        foreach (WallType wt in wallCollector)
                        {
                            types.Add(wt.Name);
                        }
                    }

                    else if (revitFamily.Category == "Floors")
                    {
                        // Get floor types
                        FilteredElementCollector floorCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        floorCollector.OfClass(typeof(FloorType));
                        floorCollector.OfCategory(BuiltInCategory.OST_Floors);
                        foreach (FloorType ft in floorCollector)
                        {
                            types.Add(ft.Name);
                        }
                    }

                    else if (revitFamily.Category == "Roofs")
                    {
                        // Get roof types
                        FilteredElementCollector roofCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        roofCollector.OfClass(typeof(RoofType));
                        roofCollector.OfCategory(BuiltInCategory.OST_Roofs);
                        foreach (RoofType rt in roofCollector)
                        {
                            types.Add(rt.Name);
                        }
                    }
                    else
                    {
                        // Get typical family types.
                        FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
                        familyCollector.OfClass(typeof(Family));
                        foreach (Family f in familyCollector)
                        {
                            if (f.Name == revitFamily.FamilyName)
                            {
                                FamilySymbolSet fss = f.Symbols;
                                foreach (FamilySymbol fs in fss)
                                {
                                    types.Add(fs.Name);
                                }

                                break;
                            }
                        }
                    }
                    typeNames = types;
                }
                catch
                {
                }
            Monitor.Wait(_locker, WAIT_TIMEOUT);
            }
            return typeNames;
        }

        public List<RevitParameter> GetParameters(RevitObject revitFamily, string typeName)
        {
            lock (_locker)
            {
                TaskContainer.Instance.EnqueueTask(uiApp =>
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    parameters = new List<RevitParameter>();
                    if (revitFamily.Category == "Walls")
                    {
                        // do stuff for walls
                        FilteredElementCollector wallCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        wallCollector.OfClass(typeof(WallType));
                        wallCollector.OfCategory(BuiltInCategory.OST_Walls);
                        foreach (WallType wt in wallCollector)
                        {
                            if (wt.Name == typeName)
                            {
                                // Get the type parameters
                                List<Parameter> typeParams = new List<Parameter>();
                                foreach (Parameter p in wt.Parameters)
                                {
                                    typeParams.Add(p);
                                }

                                // Get the instance parameters
                                List<Parameter> instParameters = new List<Parameter>();
                                using (Transaction t = new Transaction(doc, "temp family"))
                                {
                                    t.Start();
                                    Wall wall = null;
                                    try
                                    {
                                        Curve c = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 0, 0));
                                        FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                        Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                        wall = Wall.Create(doc, c, l.Id, false);
                                    }
                                    catch
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                    }
                                    if (wall != null)
                                    {
                                        foreach (Parameter p in wall.Parameters)
                                        {
                                            instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = true;
                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = false;
                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.Category == "Floors")
                    {
                        // get parameters for floors
                        FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
                        floorCollector.OfClass(typeof(FloorType));
                        floorCollector.OfCategory(BuiltInCategory.OST_Floors);
                        foreach (FloorType ft in floorCollector)
                        {
                            if (ft.Name == typeName)
                            {
                                // Get the type parameters
                                List<Parameter> typeParams = new List<Parameter>();
                                foreach (Parameter p in ft.Parameters)
                                {
                                    typeParams.Add(p);
                                }

                                // Get the instance parameters
                                List<Parameter> instParameters = new List<Parameter>();
                                using (Transaction t = new Transaction(doc, "temp family"))
                                {
                                    t.Start();
                                    Floor floor = null;
                                    try
                                    {
                                        Curve c1 = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 0, 0));
                                        Curve c2 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(1, 1, 0));
                                        Curve c3 = Line.CreateBound(new XYZ(1, 1, 0), new XYZ(0, 1, 0));
                                        Curve c4 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(0, 0, 0));
                                        CurveArray profile = new CurveArray();
                                        profile.Append(c1);
                                        profile.Append(c2);
                                        profile.Append(c3);
                                        profile.Append(c4);
                                        floor = doc.Create.NewFloor(profile, false);
                                    }
                                    catch
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                    }
                                    if (floor != null)
                                    {
                                        foreach (Parameter p in floor.Parameters)
                                        {
                                            instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = true;
                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = false;
                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.Category == "Roofs")
                    {
                        // get parameters for a roof
                        FilteredElementCollector roofCollector = new FilteredElementCollector(doc);
                        roofCollector.OfClass(typeof(RoofType));
                        roofCollector.OfCategory(BuiltInCategory.OST_Roofs);
                        foreach (RoofType rt in roofCollector)
                        {
                            if (rt.Name == typeName)
                            {
                                // Get the type parameters
                                List<Parameter> typeParams = new List<Parameter>();
                                foreach (Parameter p in rt.Parameters)
                                {
                                    typeParams.Add(p);
                                }

                                // Get the instance parameters
                                List<Parameter> instParameters = new List<Parameter>();
                                using (Transaction t = new Transaction(doc, "temp family"))
                                {
                                    t.Start();
                                    FootPrintRoof roof = null;
                                    try
                                    {
                                        Curve c1 = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 0, 0));
                                        Curve c2 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(1, 1, 0));
                                        Curve c3 = Line.CreateBound(new XYZ(1, 1, 0), new XYZ(0, 1, 0));
                                        Curve c4 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(0, 0, 0));
                                        CurveArray profile = new CurveArray();
                                        profile.Append(c1);
                                        profile.Append(c2);
                                        profile.Append(c3);
                                        profile.Append(c4);
                                        FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                        Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                        ModelCurveArray curveArrayMapping = new ModelCurveArray();
                                        roof = doc.Create.NewFootPrintRoof(profile, l, rt, out curveArrayMapping);
                                    }
                                    catch
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                    }
                                    if (roof != null)
                                    {
                                        foreach (Parameter p in roof.Parameters)
                                        {
                                            instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = true;
                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter();
                                    rp.ParameterName = p.Definition.Name;
                                    rp.StorageType = p.StorageType.ToString();
                                    rp.IsType = false;
                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else
                    {
                        // Regular family.  Proceed to get all parameters
                        // TODO: There are different options for different hostings.  
                        //       Make sure you account for this while creating the instance to get the inst parameters.
                        FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
                        familyCollector.OfClass(typeof(Family));
                        foreach (Family f in familyCollector)
                        {
                            if (f.Name == revitFamily.FamilyName)
                            {
                                FamilySymbolSet fss = f.Symbols;
                                foreach (FamilySymbol fs in fss)
                                {
                                    if (fs.Name == typeName)
                                    {
                                        List<Parameter> typeParams = new List<Parameter>();
                                        foreach (Parameter p in fs.Parameters)
                                        {
                                            typeParams.Add(p);
                                        }
                                        List<Parameter> instanceParams = new List<Parameter>();
                                        // temporary create an instance of the family to get instance parameters
                                        using (Transaction t = new Transaction(doc, "temp family"))
                                        {
                                            t.Start();
                                            FamilyInstance fi = null;
                                            // Get the hosting type
                                            int hostType = f.get_Parameter(BuiltInParameter.FAMILY_HOSTING_BEHAVIOR).AsInteger();
                                            if (hostType == 0)
                                            {
                                                // Typical
                                            }
                                            else if (hostType == 1)
                                            {
                                                // Wall hosted
                                                // Temporary wall
                                                Wall wall = null;
                                                Curve c = Line.CreateBound(new XYZ(-20, 0, 0), new XYZ(20, 0, 0));
                                                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                                Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                                try
                                                {
                                                    wall = Wall.Create(doc, c, l.Id, false);
                                                }
                                                catch
                                                {
                                                    // Failed to create the wall, no instance parameters will be found
                                                }
                                                if (wall != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fs, wall as Element, l, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                else
                                                {
                                                    // regular creation.  SOme parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                //fi = doc.Create.NewFamilyInstance(origin, fs, host, level, false)
                                            }
                                            else if (hostType == 2)
                                            {
                                                // Floor Hosted
                                                // temporary floor
                                                Floor floor = null;
                                                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                                Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                                try
                                                {
                                                    Curve c1 = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 0, 0));
                                                    Curve c2 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(1, 1, 0));
                                                    Curve c3 = Line.CreateBound(new XYZ(1, 1, 0), new XYZ(0, 1, 0));
                                                    Curve c4 = Line.CreateBound(new XYZ(0, 1, 0), new XYZ(0, 0, 0));
                                                    CurveArray profile = new CurveArray();
                                                    profile.Append(c1);
                                                    profile.Append(c2);
                                                    profile.Append(c3);
                                                    profile.Append(c4);
                                                    floor = doc.Create.NewFloor(profile, false);
                                                }
                                                catch
                                                {
                                                    // Failed to create the wall, no instance parameters will be found
                                                }
                                                if (floor != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fs, floor as Element, l, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                else
                                                {
                                                    // regular creation.  SOme parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                            }
                                            else if (hostType == 3)
                                            {
                                                // Ceiling Hosted (might be difficult)
                                            }
                                            else if (hostType == 4)
                                            {
                                                // Roof Hosted
                                            }
                                            else if (hostType == 5)
                                            {
                                                // Face Based.
                                            }
                                            // Create a typical family instance
                                            try
                                            {
                                                fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                            // TODO: Try creating other family instances like walls, sketch based, ... and getting the instance params
                                            foreach (Parameter p in fi.Parameters)
                                            {
                                                instanceParams.Add(p);
                                            }
                                            t.RollBack();
                                        }

                                        typeParams.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                        instanceParams.Sort((x, y) => String.Compare(x.Definition.Name, y.Definition.Name));
                                        foreach (Parameter p in typeParams)
                                        {
                                            RevitParameter rp = new RevitParameter();
                                            rp.ParameterName = p.Definition.Name;
                                            rp.StorageType = p.StorageType.ToString();
                                            rp.IsType = true;
                                            parameters.Add(rp);
                                        }
                                        foreach (Parameter p in instanceParams)
                                        {
                                            RevitParameter rp = new RevitParameter();
                                            rp.ParameterName = p.Definition.Name;
                                            rp.StorageType = p.StorageType.ToString();
                                            rp.IsType = false;
                                            parameters.Add(rp);
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                });
            }
            return parameters;
        }

        //public bool CreateObjects(List<RevitObject> objects)
        //{
        //    TaskContainer.Instance.EnqueueTask(uiApp =>
        //        {
        //            try
        //            {
        //                TaskDialog dlg = new TaskDialog("Warning");
        //                dlg.MainInstruction = "Incoming Data";
        //                dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
        //                    "  Do you want to accept the incoming data and create new elements?";
        //                dlg.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

        //                TaskDialogResult result = dlg.Show();
        //                if (result == TaskDialogResult.Yes)
        //                {
        //                    // Go ahead and create some new elements
        //                    //TaskDialog.Show("Temporary", "go ahead and create " + objects.Count.ToString() + " new elements");
        //                    Document doc = uiApp.ActiveUIDocument.Document;

        //                    // Find the FamilySymbol
        //                    FamilySymbol symbol = null;
        //                    FilteredElementCollector famCollector = new FilteredElementCollector(doc);
        //                    famCollector.OfClass(typeof(Family));
        //                    foreach (Family f in famCollector)
        //                    {
        //                        if (f.Name == objects[0].FamilyName)
        //                        {
        //                            foreach (FamilySymbol fs in f.Symbols)
        //                            {
        //                                if (fs.Name == objects[0].TypeName)
        //                                {
        //                                    symbol = fs;
        //                                }
        //                            }
        //                        }
        //                    }

        //                    if (symbol != null)
        //                    {
        //                        using (Transaction t = new Transaction(doc, "Create Objects"))
        //                        {
        //                            t.Start();
        //                            foreach (RevitObject ro in objects)
        //                            {
        //                                XYZ origin = new XYZ(ro.Origin.X, ro.Origin.Y, ro.Origin.Z);
        //                                FamilyInstance fi = doc.Create.NewFamilyInstance(origin, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        //                                foreach (RevitParameter rp in ro.Parameters)
        //                                {
        //                                    try
        //                                    {
        //                                        Parameter p = fi.get_Parameter(rp.ParameterName);
        //                                        switch (rp.StorageType)
        //                                        {
        //                                            case "Double":
        //                                                p.Set(Convert.ToDouble(rp.Value));
        //                                                break;
        //                                            case "Integer":
        //                                                p.Set(Convert.ToInt32(rp.Value));
        //                                                break;
        //                                            case "String":
        //                                                p.Set(rp.Value);
        //                                                break;
        //                                            case "ElementId":
        //                                                p.Set(new ElementId(Convert.ToInt32(rp.Value)));
        //                                                break;
        //                                            default:
        //                                                p.Set(rp.Value);
        //                                                break;
        //                                        }
        //                                    }
        //                                    catch { }
        //                                }
        //                                // Rotate
        //                                if (ro.Orientation != null)
        //                                {
        //                                    if (ro.Orientation.Z == 0)
        //                                    {
        //                                        Line axis = Line.CreateBound(origin, origin + XYZ.BasisZ);
        //                                        double angle = Math.Atan2(ro.Orientation.Y, ro.Orientation.X);
        //                                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);
        //                                    }
        //                                }
        //                            }
        //                            t.Commit();
        //                        }
        //                    }
        //                    else
        //                    {
        //                        TaskDialog.Show("Error", "Could not find the desired family type");
        //                    }
        //                }
        //            }
        //            finally
        //            {
        //                Monitor.Pulse(_locker);
        //            }
        //        });
        //    Monitor.Wait(_locker, WAIT_TIMEOUT);
            
        //    return true;
        //}

        public bool CreateOrModify(List<RevitObject> incomingObjs, Guid uniqueId)
        {

            TaskContainer.Instance.EnqueueTask(uiApp =>
            {
                try
                {
                    // Find existing elements
                    List<ElementId> existing = FindExisting(uiApp.ActiveUIDocument.Document, uniqueId);
                    
                    TaskDialog dlg = new TaskDialog("Warning");
                    dlg.MainInstruction = "Incoming Data";

                    int option = 0;
                    if (existing == null || existing.Count == 0)
                    {
                        option = 0;
                        dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
                            " This data will be used to create " + incomingObjs.Count.ToString() + " elements.  How would you like to proceed?";
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Create new elements");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancel");
                    }
                    else if (existing != null && existing.Count == incomingObjs.Count)
                    {
                        option = 1;
                        dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
                            " This incoming data matches up with " + incomingObjs.Count.ToString() + " elements.  How would you like to proceed?";
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Modify existing elements with incoming data");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Create new elements, ignore all exisiting.");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
                    }
                    else if (existing != null && existing.Count < incomingObjs.Count)
                    {
                        option = 2;
                        dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
                            " This incoming data matches up with " + incomingObjs.Count.ToString() + " elements but includes " + (incomingObjs.Count - existing.Count).ToString() +
                            " additional elements.  How would you like to proceed?";
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Modify existing and create new elements with incoming data");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Create new elements, ignore all exisiting.");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
                    }
                    else if (existing != null && existing.Count > incomingObjs.Count)
                    {
                        option = 3;
                        dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
                            " This incoming data matches up with " + incomingObjs.Count.ToString() + " elements but there are an additional " + (existing.Count - incomingObjs.Count).ToString() +
                            " in the model in addition to what's coming in.  How would you like to proceed?";
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Modify the first set of existing and delete additional elements to match incoming data");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Modify the first set of existing objects and ignore any additional elements.");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Create new elements, ignore all exisiting.");
                        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Cancel");
                    }

                    TaskDialogResult result = dlg.Show();
                    if (result == TaskDialogResult.CommandLink1)
                    {
                        if (option == 0)
                        {
                            // Create new
                            try
                            {
                                CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId);
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Error", ex.Message);
                            }
                        }
                        else if (option == 1)
                        {
                            // Modify
                            try
                            {
                                ModifyObjects(incomingObjs, existing, uiApp.ActiveUIDocument.Document);
                            }
                            catch { }
                        }
                        else if (option == 2)
                        {
                            // create and modify
                            List<RevitObject> existingObjects = new List<RevitObject>();
                            List<RevitObject> newObjects = new List<RevitObject>();

                            int i = 0;
                            while (i < existing.Count)
                            {
                                existingObjects.Add(incomingObjs[i]);
                                i++;
                            }
                            while (i < incomingObjs.Count)
                            {
                                newObjects.Add(incomingObjs[i]);
                                i++;
                            }
                            try
                            {
                                ModifyObjects(existingObjects, existing, uiApp.ActiveUIDocument.Document);
                                CreateObjects(newObjects, uiApp.ActiveUIDocument.Document, uniqueId);
                            }
                            catch { }
                        }
                        else if (option == 3)
                        {
                            // Modify and Delete
                            List<RevitObject> existingObjects = new List<RevitObject>();
                            List<ElementId> removeObjects = new List<ElementId>();

                            int i = 0;
                            while (i < incomingObjs.Count)
                            {
                                existingObjects.Add(incomingObjs[i]);
                                i++;
                            }
                            while (i < existing.Count)
                            {
                                removeObjects.Add(existing[i]);
                                i++;
                            }
                            try
                            {
                                ModifyObjects(existingObjects, existing, uiApp.ActiveUIDocument.Document);
                                DeleteExisting(uiApp.ActiveUIDocument.Document, removeObjects);
                            }
                            catch { }
                        }
                    }
                    else if (result == TaskDialogResult.CommandLink2)
                    {
                        // if 0, do nothing
                        if (option == 1)
                        {
                            // Create new
                            try
                            {
                                CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId);
                            }
                            catch { }
                        }
                        else if (option == 2)
                        {
                            // Create new
                            try
                            {
                                CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId);
                            }
                            catch { }
                        }
                        else if (option == 3)
                        {
                            // Modify and ignore
                            List<RevitObject> existingObjects = new List<RevitObject>();
                            List<ElementId> removeObjects = new List<ElementId>();

                            int i = 0;
                            while (i < incomingObjs.Count)
                            {
                                existingObjects.Add(incomingObjs[i]);
                                i++;
                            }
                            try
                            {
                                ModifyObjects(existingObjects, existing, uiApp.ActiveUIDocument.Document);
                            }
                            catch { }
                        }
                    }
                    else if (result == TaskDialogResult.CommandLink3)
                    {
                        if (option == 3)
                        {
                            // Create new
                            try
                            {
                                CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId);
                            }
                            catch { }
                        }
                        
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    Monitor.Pulse(_locker);
                }
            });
            Monitor.Wait(_locker, WAIT_TIMEOUT);
            
            return true;
        }

        public string GetDocumentName()
        {
            lock (_locker)
            {
                UIApplication uiapp = LMNA.Lyrebird.RevitServerApp.UIApp;
                string docName = "Not finished";
                try
                {
                    docName = uiapp.ActiveUIDocument.Document.Title;
                    currentDocName = docName;
                }
                catch
                {
                    currentDocName = "I FAILED";
                }
                Monitor.Wait(_locker, WAIT_TIMEOUT);
            }
            return currentDocName;
        }

        private void CreateObjects(List<RevitObject> revitObjects, Document doc, Guid uniqueId)
        {
            // Create new Revit objects.
            List<LyrebirdId> newUniqueIds = new List<LyrebirdId>();
            
            // Determine what kind of object we're creating.
            RevitObject ro = revitObjects[0];
            if (ro.Origin != null)
            {
                // Create normal objects
                // Find the FamilySymbol
                FamilySymbol symbol = null;
                FilteredElementCollector famCollector = new FilteredElementCollector(doc);
                famCollector.OfClass(typeof(Family));
                foreach (Family f in famCollector)
                {
                    if (f.Name == revitObjects[0].FamilyName)
                    {
                        foreach (FamilySymbol fs in f.Symbols)
                        {
                            if (fs.Name == revitObjects[0].TypeName)
                            {
                                symbol = fs;
                            }
                        }
                    }
                }

                if (symbol != null)
                {
                    using (Transaction t = new Transaction(doc, "Lyrebird Create Objects"))
                    {
                        t.Start();
                        try
                        {
                            // Create the Schema for the instances to store the GH Component InstanceGUID and the path
                            Schema instanceSchema = null;
                            try
                            {
                                instanceSchema = Schema.Lookup(instanceSchemaGUID);
                            }
                            catch { }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");

                                sb.SetSchemaName("LMNtsInstanceGUID");
                                instanceSchema = sb.Finish();
                            }
                            FamilyInstance fi = null;
                            XYZ origin = XYZ.Zero;
                            foreach (RevitObject obj in revitObjects)
                            {
                                try
                                {
                                    origin = new XYZ(obj.Origin.X, obj.Origin.Y, obj.Origin.Z);
                                    fi = doc.Create.NewFamilyInstance(origin, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    foreach (RevitParameter rp in obj.Parameters)
                                    {
                                        try
                                        {
                                            Parameter p = fi.get_Parameter(rp.ParameterName);
                                            switch (rp.StorageType)
                                            {
                                                case "Double":
                                                    p.Set(Convert.ToDouble(rp.Value));
                                                    break;
                                                case "Integer":
                                                    p.Set(Convert.ToInt32(rp.Value));
                                                    break;
                                                case "String":
                                                    p.Set(rp.Value);
                                                    break;
                                                case "ElementId":
                                                    p.Set(new ElementId(Convert.ToInt32(rp.Value)));
                                                    break;
                                                default:
                                                    p.Set(rp.Value);
                                                    break;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            TaskDialog.Show("Error", ex.Message);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Error", ex.Message);
                                }
                                // Rotate
                                if (ro.Orientation != null)
                                {
                                    if (ro.Orientation.Z == 0)
                                    {
                                        Line axis = Line.CreateBound(origin, origin + XYZ.BasisZ);
                                        double angle = Math.Atan2(ro.Orientation.Y, ro.Orientation.X);
                                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);
                                    }
                                }
                                // Assign the GH InstanceGuid
                                try
                                {
                                    Entity entity = new Entity(instanceSchema);
                                    Field field = instanceSchema.GetField("InstanceID");
                                    entity.Set<string>(field, uniqueId.ToString());
                                    fi.SetEntity(entity);
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Error", ex.Message);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", ex.Message);
                        }
                        t.Commit();
                    }
                }
                else
                {
                    TaskDialog.Show("Error", "Could not find the desired family type");
                }

            }
            else if (ro.AdaptivePoints != null && ro.AdaptivePoints.Count > 0)
            {
                // Create Adaptive objects
            }
            else if (ro.CurvePoints != null && ro.CurvePoints.Count == 1)
            {
                // single curve based element.
                if (ro.CurvePoints[0].Count == 2)
                {
                    // Check category to know what to make.
                    if (ro.Category == "Walls")
                    {
                        // create a line based wall

                    }

                    else
                    {
                        // create something line based.  Could be a beam, column, or line based family
                        // Test that the Location is a LocationPoint or a LocationCurve.
                        // LocationCurve is a line based family, if it's LocationPoint then make sure it's a structural column

                    }
                }
                else if (ro.CurvePoints[0].Count == 3)
                {
                    if (ro.Category == "Walls")
                    {
                        // Create line based wall
                    }
                    else
                    {
                        // Beam family.
                    }
                }
            }
            else if (ro.CurvePoints != null && ro.CurvePoints.Count > 2)
            {
                // Profile based creation.  Floor, wall, or roof.
                if (ro.Category == "Walls")
                {
                    // Create a profile based wall
                }
                else if (ro.Category == "Floors")
                {
                    // Create a floor

                }
                else if (ro.Category == "Roofs")
                {
                    // create a roof
                }
            }

            //return uniqueIds;
        }

        private bool ModifyObjects(List<RevitObject> existingObjects, List<ElementId> existingElems, Document doc)
        {
            bool succeeded = true;

            // TODO: Modify the existingElems according to data passed from existingObjects

            return succeeded;
        }

        private List<ElementId> FindExisting(Document doc, Guid uniqueId)
        {
            // Find existing elements with a matching GUID from the GH component.
            List<ElementId> existingElems = new List<ElementId>();
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilyInstance));
            Schema instanceSchema = Schema.Lookup(instanceSchemaGUID);
            if(instanceSchema == null)
            {
                return existingElems;
            }
            foreach (Element e in collector)
            {
                try
                {
                    FamilyInstance fi = e as FamilyInstance;
                    Entity entity = fi.GetEntity(instanceSchema);
                    if (entity.IsValid())
                    {
                        Field f = instanceSchema.GetField("InstanceID");
                        string tempId = entity.Get<string>(f);
                        if (tempId == uniqueId.ToString())
                        {
                            existingElems.Add(e.Id);
                        }
                    }
                }
                catch { }
            }


            return existingElems;
        }

        private void DeleteExisting(Document doc, List<ElementId> elements)
        {
            using (Transaction t = new Transaction(doc, "Lyrebird Delete Existing"))
            {
                t.Start();
                doc.Delete(elements);
                t.Commit();
            }
        }
    }
}
