using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage;
using LMNA.Lyrebird.LyrebirdCommon;


namespace LMNA.Lyrebird
{
    [ServiceBehavior]
    public class LyrebirdService : ILyrebirdService
    {
        private string currentDocName = "NULL";
        private List<RevitObject> familyNames = new List<RevitObject>();
        private List<string> typeNames = new List<string>();
        private List<RevitParameter> parameters = new List<RevitParameter>();
        private List<string> elementIdCollection = new List<string>();

        private int modBehavior;
        private int runId;

        public int ModifyBehavior
        {
            get { return modBehavior; }
            set { modBehavior = value; }
        }

        public int RunId
        {
            get { return runId; }
            set { runId = value; }
        }

        DisplayUnitType lengthDUT;
        DisplayUnitType areaDUT;
        DisplayUnitType volumeDUT;

        FamilyInstance hostFinder;

        private readonly Guid instanceSchemaGUID = new Guid("9ab787e0-1660-40b7-9453-94e1043b58db");

        private static readonly object _locker = new object();

        private const int WAIT_TIMEOUT = 1000;

        List<RevitObject> ILyrebirdService.GetFamilyNames()
        {
            familyNames.Add(new RevitObject("NULL", -1, "NULL"));
            lock (_locker)
            {
                try
                {
                    UIApplication uiApp = RevitServerApp.UIApp;
                    familyNames = new List<RevitObject>();
                    
                    // Get all standard wall families
                    FilteredElementCollector familyCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    familyCollector.OfClass(typeof(Family));
                    List<RevitObject> families = new List<RevitObject>();
                    foreach (Family f in familyCollector)
                    {
                        RevitObject ro = new RevitObject(f.FamilyCategory.Name, f.FamilyCategory.Id.IntegerValue, f.Name);
                        families.Add(ro);
                    }
                    
                    // Add System families
                    FilteredElementCollector wallTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    wallTypeCollector.OfClass(typeof(WallType));
                    WallType wt = wallTypeCollector.FirstElement() as WallType;
                    if (wt != null)
                    {
                        RevitObject wallObj = new RevitObject(wt.Category.Name, wt.Category.Id.IntegerValue, wt.Category.Name);
                        families.Add(wallObj);
                    }
                    
                    //RevitObject curtainObj = new RevitObject("Walls", "Curtain Wall");
                    //families.Add(curtainObj);
                    //RevitObject stackedObj = new RevitObject("Walls", "Stacked Wall");
                    //families.Add(stackedObj);

                    FilteredElementCollector floorTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    floorTypeCollector.OfClass(typeof(FloorType));
                    FloorType ft = floorTypeCollector.FirstElement() as FloorType;
                    if (ft != null)
                    {
                        RevitObject floorObj = new RevitObject(ft.Category.Name, ft.Category.Id.IntegerValue, ft.Category.Name);
                        families.Add(floorObj);
                    }
                    
                    FilteredElementCollector roofTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    roofTypeCollector.OfClass(typeof(RoofType));
                    RoofType rt = roofTypeCollector.FirstElement() as RoofType;
                    if (rt != null)
                    {
                        RevitObject roofObj = new RevitObject(rt.Category.Name, rt.Category.Id.IntegerValue, rt.Category.Name);
                        families.Add(roofObj);
                    }
                    
                    FilteredElementCollector levelTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    levelTypeCollector.OfClass(typeof(LevelType));
                    LevelType lt = levelTypeCollector.FirstElement() as LevelType;
                    if (lt != null)
                    {
                        RevitObject levelObj = new RevitObject(lt.Category.Name, lt.Category.Id.IntegerValue, lt.Category.Name);
                        families.Add(levelObj);
                    }
                    
                    FilteredElementCollector gridTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    gridTypeCollector.OfClass(typeof(GridType));
                    GridType gt = gridTypeCollector.FirstElement() as GridType;
                    if (gt != null)
                    {
                        RevitObject gridObj = new RevitObject(gt.Category.Name, gt.Category.Id.IntegerValue, gt.Category.Name);
                        families.Add(gridObj);
                    }
                    
                    FilteredElementCollector lineTypeCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                    lineTypeCollector.OfCategory(BuiltInCategory.OST_Lines);
                    try
                    {
                        CurveElement ce = lineTypeCollector.FirstElement() as CurveElement;
                        RevitObject modelLineObj = new RevitObject(ce.Category.Name, ce.Category.Id.IntegerValue, "Model Lines");
                        RevitObject detailLineObj = new RevitObject(ce.Category.Name, ce.Category.Id.IntegerValue, "Detail Lines");
                        families.Add(modelLineObj);
                        families.Add(detailLineObj);
                    }
                    catch
                    {
                        RevitObject modelLineObj = new RevitObject("Lines", -2000051, "Model Lines");
                        RevitObject detailLineObj = new RevitObject("Lines", -2000051, "Detail Lines");
                        families.Add(modelLineObj);
                        families.Add(detailLineObj);
                    }
                    families.Sort((x, y) => String.CompareOrdinal(x.FamilyName.ToUpper(), y.FamilyName.ToUpper()));
                    familyNames = families;
                }
                catch (Exception exception)
                {
                    Debug.WriteLine(exception.Message);
                }
                Monitor.Wait(_locker, Properties.Settings.Default.infoTimeout);
            }
            return familyNames;
        }

        List<string> ILyrebirdService.GetTypeNames(RevitObject revitFamily)
        {
            typeNames.Add("NULL");
            lock (_locker)
            {
                try
                {
                    UIApplication uiApp = RevitServerApp.UIApp;
                    var doc = uiApp.ActiveUIDocument.Document;
                    typeNames = new List<string>();
                    List<string> types = new List<string>();
                    if (revitFamily.CategoryId == -2000011)
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

                    else if (revitFamily.CategoryId == -2000032)
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

                    else if (revitFamily.CategoryId == -2000035)
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
                    else if (revitFamily.CategoryId == -2000240)
                    {
                        // Get Level Types
                        FilteredElementCollector levelCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        levelCollector.OfClass(typeof(LevelType));
                        levelCollector.OfCategory(BuiltInCategory.OST_Levels);
                        foreach (LevelType lt in levelCollector)
                        {
                            types.Add(lt.Name);
                        }
                    }
                    else if (revitFamily.CategoryId == -2000220)
                    {
                        // Get Grid Types
                        FilteredElementCollector gridCollector = new FilteredElementCollector(uiApp.ActiveUIDocument.Document);
                        gridCollector.OfClass(typeof(GridType));
                        gridCollector.OfCategory(BuiltInCategory.OST_Grids);
                        foreach (GridType gt in gridCollector)
                        {
                            types.Add(gt.Name);
                        }
                    }
                    else if (revitFamily.CategoryId == -2000051)
                    {
                        Categories cats = doc.Settings.Categories;
                        Category lineCat = null;
                        foreach (Category cat in cats)
                        {
                            if (cat.Id.IntegerValue == -2000051)
                            {
                                lineCat = cat;
                                break;
                            }
                        }

                        if (lineCat != null)
                        {
                            foreach (Category subCat in lineCat.SubCategories)
                            {
                                if (revitFamily.FamilyName == "Model Lines")
                                    types.Add(subCat.Name);
                                else if (revitFamily.FamilyName == "Detail Lines" && !subCat.Name.StartsWith("<") && !subCat.Name.EndsWith(">"))
                                    types.Add(subCat.Name);

                            }
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
                                ISet<ElementId> fsIds = f.GetFamilySymbolIds();
                                foreach (ElementId fsid in fsIds)
                                {
                                    FamilySymbol fs = doc.GetElement(fsid) as FamilySymbol;
                                    types.Add(fs.Name);
                                }

                                break;
                            }
                        }
                    }
                    types.Sort((x, y) => String.CompareOrdinal(x.ToUpper(), y.ToUpper()));
                    typeNames = types;
                }
                catch (Exception exception)
                {
                    Debug.WriteLine(exception.Message);
                }
                Monitor.Wait(_locker, Properties.Settings.Default.infoTimeout);
            }
            return typeNames;
        }

        List<RevitParameter> ILyrebirdService.GetParameters(RevitObject revitFamily, string typeName)
        {
            lock (_locker)
            {
                TaskContainer.Instance.EnqueueTask(uiApp =>
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    parameters = new List<RevitParameter>();
                    if (revitFamily.CategoryId == -2000011)
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
                                    if(!p.IsReadOnly)
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
                                        if (l != null) wall = Wall.Create(doc, c, l.Id, false);
                                    }
                                    catch (Exception exception)
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                        Debug.WriteLine(exception.Message);
                                    }

                                    if (wall != null)
                                    {
                                        foreach (Parameter p in wall.Parameters)
                                        {
                                            //if(!p.IsReadOnly)
                                                instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = true
                                    };
                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = false
                                    };

                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.CategoryId == -2000032)
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
                                    if(!p.IsReadOnly)
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

                                    catch (Exception ex)
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                        Debug.WriteLine(ex.Message);
                                    }
                                    if (floor != null)
                                    {
                                        foreach (Parameter p in floor.Parameters)
                                        {
                                            if (!p.IsReadOnly)
                                                instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = true
                                    };

                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = false
                                    };

                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.CategoryId == -2000035)
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
                                    if (!p.IsReadOnly)
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
                                    catch (Exception ex)
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                        Debug.WriteLine(ex.Message);
                                    }
                                    if (roof != null)
                                    {
                                        foreach (Parameter p in roof.Parameters)
                                        {
                                            if (!p.IsReadOnly)
                                                instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = true
                                    };

                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = false
                                    };

                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.CategoryId == -2000240)
                    {
                        // Level Families
                        FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                        levelCollector.OfClass(typeof(LevelType));
                        levelCollector.OfCategory(BuiltInCategory.OST_Levels);
                        foreach (LevelType lt in levelCollector)
                        {
                            if (lt.Name == typeName)
                            {
                                // Get the type parameters
                                List<Parameter> typeParams = new List<Parameter>();
                                foreach (Parameter p in lt.Parameters)
                                {
                                    if (!p.IsReadOnly)
                                        typeParams.Add(p);
                                }

                                // Get the instance parameters
                                List<Parameter> instParameters = new List<Parameter>();
                                using (Transaction t = new Transaction(doc, "temp level"))
                                {
                                    t.Start();
                                    Level lvl = null;
                                    try
                                    {
                                        lvl = doc.Create.NewLevel(-1000.22);
                                    }
                                    catch (Exception ex)
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                        Debug.WriteLine(ex.Message);
                                    }
                                    if (lvl != null)
                                    {
                                        foreach (Parameter p in lvl.Parameters)
                                        {
                                            if (!p.IsReadOnly)
                                                instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = true
                                    };

                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = false
                                    };

                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.CategoryId == -2000220)
                    {
                        // Grid Families
                        FilteredElementCollector gridCollector = new FilteredElementCollector(doc);
                        gridCollector.OfClass(typeof(GridType));
                        gridCollector.OfCategory(BuiltInCategory.OST_Grids);
                        foreach (GridType gt in gridCollector)
                        {
                            if (gt.Name == typeName)
                            {
                                // Get the type parameters
                                List<Parameter> typeParams = new List<Parameter>();
                                foreach (Parameter p in gt.Parameters)
                                {
                                    if (!p.IsReadOnly)
                                        typeParams.Add(p);
                                }

                                // Get the instance parameters
                                List<Parameter> instParameters = new List<Parameter>();
                                using (Transaction t = new Transaction(doc, "temp grid"))
                                {
                                    t.Start();
                                    Grid grid = null;
                                    try
                                    {
                                        Line ln = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 1, 0));
                                        grid = doc.Create.NewGrid(ln);
                                    }
                                    catch (Exception ex)
                                    {
                                        // Failed to create the wall, no instance parameters will be found
                                        Debug.WriteLine(ex.Message);
                                    }
                                    if (grid != null)
                                    {
                                        foreach (Parameter p in grid.Parameters)
                                        {
                                            if (!p.IsReadOnly)
                                                instParameters.Add(p);
                                        }
                                    }
                                    t.RollBack();
                                }
                                typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                instParameters.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                foreach (Parameter p in typeParams)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = true
                                    };

                                    parameters.Add(rp);
                                }
                                foreach (Parameter p in instParameters)
                                {
                                    RevitParameter rp = new RevitParameter
                                    {
                                        ParameterName = p.Definition.Name,
                                        StorageType = p.StorageType.ToString(),
                                        IsType = false
                                    };

                                    parameters.Add(rp);
                                }
                                break;
                            }
                        }
                    }
                    else if (revitFamily.CategoryId == -2000051)
                    {
                        // leave parameters empty
                    }
                    else
                    {
                        // Regular family.  Proceed to get all parameters
                        FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
                        familyCollector.OfClass(typeof(Family));
                        foreach (Family f in familyCollector)
                        {
                            if (f.Name == revitFamily.FamilyName)
                            {
                                ISet<ElementId> fsIds = f.GetFamilySymbolIds();

                                foreach (ElementId fsid in fsIds)
                                {
                                    FamilySymbol fs = doc.GetElement(fsid) as FamilySymbol;
                                    if (fs.Name == typeName)
                                    {
                                        List<Parameter> typeParams = new List<Parameter>();
                                        foreach (Parameter p in fs.Parameters)
                                        {
                                            if (!p.IsReadOnly)
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
                                                    if (l != null) wall = Wall.Create(doc, c, l.Id, false);
                                                }
                                                catch (Exception ex)
                                                {
                                                    // Failed to create the wall, no instance parameters will be found
                                                    Debug.WriteLine(ex.Message);
                                                }
                                                if (wall != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fs, wall as Element, l, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                else
                                                {
                                                    // regular creation.  Some parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
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
                                                catch (Exception ex)
                                                {
                                                    // Failed to create the floor, no instance parameters will be found
                                                    Debug.WriteLine(ex.Message);
                                                }
                                                if (floor != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fs, floor as Element, l, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                else
                                                {
                                                    // regular creation.  Some parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                            }
                                            else if (hostType == 3)
                                            {
                                                // Ceiling Hosted (might be difficult)
                                                // Try to find a ceiling
                                                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                                Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                                FilteredElementCollector ceilingCollector = new FilteredElementCollector(doc);
                                                Ceiling ceiling = ceilingCollector.OfClass(typeof(Ceiling)).ToElements().OfType<Ceiling>().FirstOrDefault();
                                                if (ceiling != null)
                                                {
                                                    // Find a point on the ceiling
                                                    Options opt = new Options();
                                                    opt.ComputeReferences = true;
                                                    GeometryElement ge = ceiling.get_Geometry(opt);
                                                    List<List<XYZ>> verticePoints = new List<List<XYZ>>();
                                                    foreach (GeometryObject go in ge)
                                                    {
                                                        Solid solid = go as Solid;
                                                        if (null == solid || 0 == solid.Faces.Size)
                                                        {
                                                            continue;
                                                        }

                                                        PlanarFace planarFace = null;
                                                        double faceArea = 0;
                                                        foreach (Face face in solid.Faces)
                                                        {

                                                            PlanarFace pf = null;
                                                            try
                                                            {
                                                                pf = face as PlanarFace;
                                                            }
                                                            catch
                                                            {
                                                            }
                                                            if (pf != null)
                                                            {
                                                                if (pf.Area > faceArea)
                                                                {
                                                                    planarFace = pf;
                                                                }
                                                            }
                                                        }
                                                        if (planarFace != null)
                                                        {
                                                            Mesh mesh = planarFace.Triangulate();
                                                            int triCnt = mesh.NumTriangles;
                                                            MeshTriangle bigTriangle = null;
                                                            for (int tri = 0; tri < triCnt; tri++)
                                                            {
                                                                if (bigTriangle == null)
                                                                {

                                                                    bigTriangle = mesh.get_Triangle(tri);
                                                                }
                                                                else
                                                                {
                                                                    MeshTriangle mt = mesh.get_Triangle(tri);
                                                                    double area = Math.Abs(((mt.get_Vertex(0).X * (mt.get_Vertex(1).Y - mt.get_Vertex(2).Y)) + (mt.get_Vertex(1).X * (mt.get_Vertex(2).Y - mt.get_Vertex(0).Y)) + (mt.get_Vertex(2).X * (mt.get_Vertex(0).Y - mt.get_Vertex(1).Y))) / 2);
                                                                    double bigTriArea = Math.Abs(((bigTriangle.get_Vertex(0).X * (bigTriangle.get_Vertex(1).Y - bigTriangle.get_Vertex(2).Y)) + (bigTriangle.get_Vertex(1).X * (bigTriangle.get_Vertex(2).Y - bigTriangle.get_Vertex(0).Y)) + (bigTriangle.get_Vertex(2).X * (bigTriangle.get_Vertex(0).Y - bigTriangle.get_Vertex(1).Y))) / 2);
                                                                    if (area > bigTriArea)
                                                                    {
                                                                        bigTriangle = mt;
                                                                    }
                                                                }
                                                            }
                                                            if (bigTriangle != null)
                                                            {
                                                                double test = Math.Abs(((bigTriangle.get_Vertex(0).X * (bigTriangle.get_Vertex(1).Y - bigTriangle.get_Vertex(2).Y)) + (bigTriangle.get_Vertex(1).X * (bigTriangle.get_Vertex(2).Y - bigTriangle.get_Vertex(0).Y)) + (bigTriangle.get_Vertex(2).X * (bigTriangle.get_Vertex(0).Y - bigTriangle.get_Vertex(1).Y))) / 2);
                                                            }
                                                            try
                                                            {
                                                                List<XYZ> ptList = new List<XYZ>();
                                                                ptList.Add(bigTriangle.get_Vertex(0));
                                                                ptList.Add(bigTriangle.get_Vertex(1));
                                                                ptList.Add(bigTriangle.get_Vertex(2));
                                                                verticePoints.Add(ptList);
                                                            }
                                                            catch
                                                            {
                                                            }
                                                            break;
                                                        }
                                                    }

                                                    if (verticePoints.Count > 0)
                                                    {
                                                        List<XYZ> vertices = verticePoints[0];
                                                        XYZ midXYZ = vertices[1] + (0.5 * (vertices[2] - vertices[1]));
                                                        XYZ centerPt = vertices[0] + (0.666667 * (midXYZ - vertices[0]));

                                                        if (ceiling != null)
                                                        {
                                                            fi = doc.Create.NewFamilyInstance(centerPt, fs, ceiling as Element, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                        }
                                                        else
                                                        {
                                                            // regular creation.  Some parameters will be missing
                                                            fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                        }

                                                    }
                                                }
                                            }
                                            else if (hostType == 4)
                                            {
                                                // Roof Hosted
                                                // Temporary roof
                                                FootPrintRoof roof = null;
                                                FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
                                                Level l = lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>().FirstOrDefault();
                                                FilteredElementCollector roofTypeCollector = new FilteredElementCollector(doc);
                                                RoofType rt = roofTypeCollector.OfClass(typeof(RoofType)).ToElements().OfType<RoofType>().FirstOrDefault();
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
                                                    ModelCurveArray roofProfile = new ModelCurveArray();
                                                    roof = doc.Create.NewFootPrintRoof(profile, l, rt, out roofProfile);
                                                }
                                                catch (Exception ex)
                                                {
                                                    // Failed to create the roof, no instance parameters will be found
                                                    Debug.WriteLine(ex.Message);
                                                }

                                                if (roof != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fs, roof as Element, l, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                                else
                                                {
                                                    // regular creation.  Some parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                            }
                                            else if (hostType == 5)
                                            {
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
                                                catch (Exception ex)
                                                {
                                                    // Failed to create the floor, no instance parameters will be found
                                                    Debug.WriteLine(ex.Message);
                                                }

                                                // Find a face on the floor to host to.
                                                Face face = FindFace(XYZ.Zero, XYZ.BasisZ, doc);
                                                if (face != null)
                                                {
                                                    fi = doc.Create.NewFamilyInstance(face, XYZ.Zero, new XYZ(0, -1, 0), fs);
                                                }
                                                else
                                                {
                                                    // regular creation.  Some parameters will be missing
                                                    fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                }
                                            }
                                            // Create a typical family instance
                                            try
                                            {
                                                fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                            }
                                            catch (Exception ex)
                                            {
                                                Debug.WriteLine(ex.Message);
                                            }
                                            // TODO: Try creating other family instances like walls, sketch based, ... and getting the instance params
                                            if (fi != null)
                                            {
                                                foreach (Parameter p in fi.Parameters)
                                                {
                                                    if (!p.IsReadOnly)
                                                        instanceParams.Add(p);
                                                }
                                            }

                                            t.RollBack();
                                        }

                                        typeParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                        instanceParams.Sort((x, y) => String.CompareOrdinal(x.Definition.Name, y.Definition.Name));
                                        foreach (Parameter p in typeParams)
                                        {
                                            RevitParameter rp = new RevitParameter
                                            {
                                                ParameterName = p.Definition.Name,
                                                StorageType = p.StorageType.ToString(),
                                                IsType = true
                                            };

                                            parameters.Add(rp);
                                        }
                                        foreach (Parameter p in instanceParams)
                                        {
                                            RevitParameter rp = new RevitParameter
                                            {
                                                ParameterName = p.Definition.Name,
                                                StorageType = p.StorageType.ToString(),
                                                IsType = false
                                            };

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

        List<string> ILyrebirdService.GetCategoryElements(ElementIdCategory eic)
        {
            elementIdCollection = new List<string>();
            elementIdCollection.Add("NULL");
            lock (_locker)
            {
                try
                {
                    UIApplication uiApp = RevitServerApp.UIApp;
                    var doc = uiApp.ActiveUIDocument.Document;
                    elementIdCollection = new List<string>();

                    if (eic == ElementIdCategory.DesignOption)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_DesignOptions);

                        foreach (Element elem in collector)
                        {
                            try
                            {
                                DesignOption desOpt = elem as DesignOption;
                                elementIdCollection.Add(desOpt.Name + "," + desOpt.Id.IntegerValue.ToString());
                            }
                            catch { }
                        }
                        elementIdCollection.Sort();
                    }

                    else if (eic == ElementIdCategory.Image)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_RasterImages);

                        foreach (Element elem in collector)
                        {
                            try
                            {
                                ImageType imgType = elem as ImageType;
                                elementIdCollection.Add(imgType.Name + "," + imgType.Id.IntegerValue.ToString());
                            }
                            catch { }
                        }
                        elementIdCollection.Sort();
                    }

                    else if (eic == ElementIdCategory.Level)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_Levels);
                        List<Level> levels = new List<Level>();
                        foreach (Element elem in collector)
                        {
                            try
                            {
                                Level lvl = elem as Level;
                                double elev = lvl.ProjectElevation;
                                levels.Add(lvl);
                            }
                            catch { }
                        }
                        
                        // Sort the levels by elevation
                        levels.Sort((lvl1, lvl2) => lvl1.Elevation.CompareTo(lvl2.Elevation));
                        
                        foreach (Level lvl in levels)
                        {
                            try
                            {
                                elementIdCollection.Add(lvl.Name + " [" + lvl.ProjectElevation.ToString() + "]," + lvl.Id.IntegerValue.ToString());
                            }
                            catch { }
                        }
                    }

                    else if (eic == ElementIdCategory.Material)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_Materials);

                        foreach (Element elem in collector)
                        {
                            try
                            {
                                Material mat = elem as Material;
                                elementIdCollection.Add(mat.Name + "," + mat.Id.IntegerValue.ToString());
                            }
                            catch { }
                        }
                        elementIdCollection.Sort();
                    }

                    else if (eic == ElementIdCategory.Phase)
                    {
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_Phases);

                        foreach (Element elem in collector)
                        {
                            try
                            {
                                Phase phase = elem as Phase;
                                elementIdCollection.Add(phase.Name + "," + phase.Id.IntegerValue.ToString());
                            }
                            catch { }
                        }
                        elementIdCollection.Sort();
                    }
                }
                catch { }
            }
            return elementIdCollection;
        }

        bool ILyrebirdService.CreateOrModify(List<RevitObject> incomingObjs, Guid uniqueId, string nickName)
        {
            lock (_locker)
            {
                TaskContainer.Instance.EnqueueTask(uiApp =>
                {
                    try
                    {
                        // Set the DisplayUnitTypes
                        Units units = uiApp.ActiveUIDocument.Document.GetUnits();
                        FormatOptions fo = units.GetFormatOptions(UnitType.UT_Length);
                        lengthDUT = fo.DisplayUnits;
                        fo = units.GetFormatOptions(UnitType.UT_Area);
                        areaDUT = fo.DisplayUnits;
                        fo = units.GetFormatOptions(UnitType.UT_Volume);
                        volumeDUT = fo.DisplayUnits;

                        // Find existing elements
                        List<ElementId> existing = FindExisting(uiApp.ActiveUIDocument.Document, uniqueId, incomingObjs[0].CategoryId, -1);
                        
                        // find if there's more than one run existing
                        Schema instanceSchema = Schema.Lookup(instanceSchemaGUID);
                        List<int> runIds = new List<int>();
                        List<Runs> allRuns = new List<Runs>();
                        if (instanceSchema != null)
                        {
                            foreach (ElementId eid in existing)
                            {
                                Element e = uiApp.ActiveUIDocument.Document.GetElement(eid);
                                // Find the run ID
                                Entity entity = e.GetEntity(instanceSchema);
                                if (entity.IsValid())
                                {
                                    Field f = instanceSchema.GetField("RunID");
                                    int tempId = entity.Get<int>(f);
                                    if (!runIds.Contains(tempId))
                                    {
                                        runIds.Add(tempId);
                                        string familyName = string.Empty;
                                        if (e.Category.Id.IntegerValue == -2000011)
                                        {
                                            Wall w = e as Wall;
                                            familyName = w.Category.Name + " : " + w.WallType.Name;
                                        }
                                        else if (e.Category.Id.IntegerValue == -2000032)
                                        {
                                            Floor flr = e as Floor;
                                            familyName = flr.Category.Name + " : " + flr.FloorType.Name;
                                        }
                                        else if (e.Category.Id.IntegerValue == -2000035)
                                        {
                                            RoofBase r = e as RoofBase;
                                            familyName = r.Category.Name + " : " + r.RoofType.Name;
                                        }
                                        else if (e.Category.Id.IntegerValue == -2000240)
                                        {
                                            Level lvl = e as Level;
                                            familyName = lvl.Category.Name + " : " + lvl.LevelType.Name;
                                        }
                                        else if (e.Category.Id.IntegerValue == -2000220)
                                        {
                                            Grid g = e as Grid;
                                            familyName = g.Category.Name + " : " + g.GridType.Name;
                                        }
                                        else
                                        {
                                            FamilyInstance famInst = e as FamilyInstance;
                                            familyName = famInst.Symbol.Family.Name + " : " + famInst.Symbol.Name;
                                        }
                                        Runs run = new Runs(tempId, "Run" + tempId.ToString(), familyName);
                                        allRuns.Add(run);
                                    }
                                }
                            }
                        }


                        if (runIds != null && runIds.Count > 0)
                        {
                            runIds.Sort((x, y) => x.CompareTo(y));
                            int lastId = 0;
                            lastId = runIds.Last();
                            ModifyForm mform = new ModifyForm(this, allRuns);
                            mform.ShowDialog();

                            // Get the set of existing elements to reflect the run choice.
                            List<ElementId> existingRunEID = FindExisting(uiApp.ActiveUIDocument.Document, uniqueId, incomingObjs[0].CategoryId, runId);

                            // modBehavior = 0, Modify the selected run
                            if (modBehavior == 0)
                            {
                                if (existingRunEID.Count == incomingObjs.Count)
                                {
                                    // just modify
                                    ModifyObjects(incomingObjs, existingRunEID, uiApp.ActiveUIDocument.Document, uniqueId, true, nickName, runId);
                                }
                                else if (existingRunEID.Count > incomingObjs.Count)
                                {
                                    // Modify and Delete
                                    List<ElementId> modObjects = new List<ElementId>();
                                    List<ElementId> removeObjects = new List<ElementId>();

                                    int i = 0;
                                    while (i < incomingObjs.Count)
                                    {
                                        modObjects.Add(existingRunEID[i]);
                                        i++;
                                    }
                                    while (existingRunEID != null && i < existingRunEID.Count)
                                    {
                                        Element e = uiApp.ActiveUIDocument.Document.GetElement(existing[i]);
                                        // Find the run ID
                                        Entity entity = e.GetEntity(instanceSchema);
                                        if (entity.IsValid())
                                        {
                                            removeObjects.Add(existingRunEID[i]);
                                            //Field f = instanceSchema.GetField("RunID");
                                            //int tempId = entity.Get<int>(f);
                                            //if (tempId == runId)
                                            //{
                                            //    removeObjects.Add(existing[i]);
                                            //}
                                        }
                                        i++;
                                    }
                                    try
                                    {
                                        ModifyObjects(incomingObjs, modObjects, uiApp.ActiveUIDocument.Document, uniqueId, true, nickName, runId);
                                        DeleteExisting(uiApp.ActiveUIDocument.Document, removeObjects);
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine(ex.Message);
                                    }
                                }
                                else if (existingRunEID.Count < incomingObjs.Count)
                                {
                                    // modify and create
                                    // create and modify
                                    List<RevitObject> existingObjects = new List<RevitObject>();
                                    List<RevitObject> newObjects = new List<RevitObject>();

                                    int i = 0;
                                    Debug.Assert(existing != null, "existing != null");
                                    while (i < existingRunEID.Count)
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
                                        ModifyObjects(existingObjects, existingRunEID, uiApp.ActiveUIDocument.Document, uniqueId, true, nickName, runId);
                                        CreateObjects(newObjects, uiApp.ActiveUIDocument.Document, uniqueId, runId, nickName);
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine(ex.Message);
                                    }
                                }
                            }
                            // modBehavior = 1, Create a new run
                            else if (modBehavior == 1)
                            {
                                // Just send everything to create a new Run.
                                try
                                {
                                    CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId, lastId + 1, nickName);
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Error", ex.Message);
                                }
                            }

                            // modBehavior = 3, cancel/ignore
                        }
                        else
                        {
                            TaskDialog dlg = new TaskDialog("Warning") { MainInstruction = "Incoming Data" };
                            RevitObject existingObj1 = incomingObjs[0];
                            bool profileWarning1 = (existingObj1.CategoryId == -2000011 && existingObj1.Curves.Count > 1) || existingObj1.CategoryId == -2000032 || existingObj1.CategoryId == -2000035;
                            if (existing == null || existing.Count == 0)
                            {
                                dlg.MainContent = "Data is being sent to Revit from another application using Lyrebird." +
                                    " This data will be used to create " + incomingObjs.Count.ToString(CultureInfo.InvariantCulture) + " elements.  How would you like to proceed?";
                                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Create new elements");
                                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancel");
                            }

                            TaskDialogResult result1 = dlg.Show();
                            if (result1 == TaskDialogResult.CommandLink1)
                            {
                                // Create new
                                try
                                {
                                    CreateObjects(incomingObjs, uiApp.ActiveUIDocument.Document, uniqueId, 0, nickName);
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Error", ex.Message);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Test", ex.Message);
                        Debug.WriteLine(ex.Message);
                    }
                    finally
                    {
                        Monitor.Pulse(_locker);
                    }
                });
                Monitor.Wait(_locker, Properties.Settings.Default.infoTimeout);
            }
            return true;
        }

        string ILyrebirdService.GetDocumentName()
        {
            lock (_locker)
            {
                UIApplication uiapp = RevitServerApp.UIApp;
                try
                {
                    currentDocName = uiapp.ActiveUIDocument.Document.Title;
                }
                catch
                {
                    currentDocName = "NULL";
                }
                Monitor.Wait(_locker, Properties.Settings.Default.infoTimeout);
            }
            return currentDocName;
        }

        private void CreateObjects(List<RevitObject> revitObjects, Document doc, Guid uniqueId, int runId, string nickName)
        {
            // Create new Revit objects.
            //List<LyrebirdId> newUniqueIds = new List<LyrebirdId>();

            // Get the levels from the project
            FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
            lvlCollector.OfClass(typeof(Level)).ToElements().OfType<Level>();

            // Determine what kind of object we're creating.
            RevitObject ro = revitObjects[0];

            #region Normal Origin based Family Instance
            if (ro.Origin != null)
            {
                // Find the FamilySymbol
                FamilySymbol symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);
                
                LevelType lt = FindLevelType(ro.TypeName, doc);

                if (symbol != null || lt != null)
                {
                    // Get the hosting ID from the family.
                    Family fam = null;
                    Parameter hostParam = null;
                    int hostBehavior = 0;

                    try
                    {
                        fam = symbol.Family;
                        hostParam = fam.get_Parameter(BuiltInParameter.FAMILY_HOSTING_BEHAVIOR);
                        hostBehavior = hostParam.AsInteger();
                    }
                    catch{}
                    
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
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");
                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");
                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }
                            FamilyInstance fi = null;
                            XYZ origin = XYZ.Zero;
                            if (lt != null)
                            {
                                // Create a level for the object.
                                foreach (RevitObject obj in revitObjects)
                                {
                                    try
                                    {
                                        Level lvl = doc.Create.NewLevel(UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT));
                                        lvl.LevelType = lt;
                                        
                                        // Set the parameters.
                                        SetParameters(lvl, obj.Parameters, doc);

                                        // Assign the GH InstanceGuid
                                        AssignGuid(lvl, uniqueId, instanceSchema, runId, nickName);
                                    }
                                    catch { }
                                }
                            }
                            else if (hostBehavior == 0)
                            {
                                int x = 0;
                                foreach (RevitObject obj in revitObjects)
                                {
                                    try
                                    {
                                        List<LyrebirdPoint> originPts = new List<LyrebirdPoint>();
                                        Level lvl = GetLevel(originPts, doc);
                                        origin = new XYZ(UnitUtils.ConvertToInternalUnits(obj.Origin.X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT));
                                        if (symbol.Category.Id.IntegerValue == -2001330)
                                        {
                                            if (lvl != null)
                                            {
                                                // Structural Column
                                                fi = doc.Create.NewFamilyInstance(origin, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.Column);
                                                fi.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(origin.Z - lvl.ProjectElevation);
                                                double topElev = ((Level)doc.GetElement(fi.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId())).ProjectElevation;
                                                if (lvl.ProjectElevation + (origin.Z - lvl.ProjectElevation) > topElev)
                                                {
                                                    fi.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set((lvl.ProjectElevation + (origin.Z - lvl.ProjectElevation)) - topElev + 10.0);
                                                }
                                            }
                                            else
                                            {
                                                TaskDialog.Show("error", "Null level");
                                            }
                                        }
                                        else
                                        {
                                            // All Else
                                            fi = doc.Create.NewFamilyInstance(origin, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", ex.Message);
                                    }

                                    // Rotate
                                    if (obj.Orientation != null)
                                    {
                                        if (Math.Round(Math.Abs(obj.Orientation.Z - 0), 10) < double.Epsilon)
                                        {
                                            Line axis = Line.CreateBound(origin, origin + XYZ.BasisZ);
                                            XYZ normalVector = new XYZ(0, -1, 0);
                                            XYZ orient = new XYZ(obj.Orientation.X, obj.Orientation.Y, obj.Orientation.Z);
                                            double angle = 0;
                                            if (orient.X < 0 && orient.Y < 0)
                                            {
                                                angle = (2 * Math.PI) - normalVector.AngleTo(orient);
                                            }
                                            else if (orient.X < 0)
                                            {
                                                angle = (Math.PI - normalVector.AngleTo(orient)) + Math.PI;
                                            }
                                            else if (orient.Y == 0)
                                            {
                                                angle = 1.5 * Math.PI;
                                            }
                                            else
                                            {
                                                angle = normalVector.AngleTo(orient);
                                            }
                                            ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);
                                        }
                                    }

                                    // Assign the parameters
                                    SetParameters(fi, obj.Parameters, doc);

                                    // Assign the GH InstanceGuid
                                    AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);

                                    x++;
                                }
                            }
                            else
                            {
                                foreach (RevitObject obj in revitObjects)
                                {
                                    origin = new XYZ(UnitUtils.ConvertToInternalUnits(obj.Origin.X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT));
                                    
                                    // Find the level
                                    List<LyrebirdPoint> lbPoints = new List<LyrebirdPoint> { obj.Origin };
                                    Level lvl = GetLevel(lbPoints, doc);
                                    
                                    // Get the host
                                    if (hostBehavior == 5)
                                    {
                                        // Face based family.  Find the face and create the element
                                        XYZ normVector = new XYZ(obj.Orientation.X, obj.Orientation.Y, obj.Orientation.Z);
                                        XYZ faceVector;
                                        if (obj.FaceOrientation != null)
                                        {
                                            faceVector = new XYZ(obj.FaceOrientation.X, obj.FaceOrientation.Y, obj.FaceOrientation.Z);
                                        }
                                        else
                                        {
                                            faceVector = XYZ.BasisZ;
                                        }
                                        Face face = FindFace(origin, normVector, doc);
                                        if (face != null)
                                        {
                                            fi = doc.Create.NewFamilyInstance(face, origin, faceVector, symbol);
                                        }
                                    }
                                    else
                                    {
                                        // typical hosted family.  Can be wall, floor, roof or ceiling.
                                        ElementId host = FindHost(origin, hostBehavior, doc);
                                        
                                        if (host != null)
                                        {
                                            fi = doc.Create.NewFamilyInstance(origin, symbol, doc.GetElement(host), lvl, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        }
                                    }
                                    // Assign the parameters
                                    SetParameters(fi, obj.Parameters, doc);

                                    // Assign the GH InstanceGuid
                                    AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                }
                                // delete the host finder
                                ElementId hostFinderFamily = hostFinder.Symbol.Family.Id;
                                doc.Delete(hostFinder.Id);
                                doc.Delete(hostFinderFamily);
                                hostFinder = null;
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
            #endregion

            #region Adaptive Components
            else if (ro.AdaptivePoints != null && ro.AdaptivePoints.Count > 0)
            {
                // Find the FamilySymbol
                FamilySymbol symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);
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
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");

                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");

                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }
                            try
                            {
                                foreach (RevitObject obj in revitObjects)
                                {
                                    FamilyInstance fi = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(doc, symbol);
                                    IList<ElementId> placePointIds = new List<ElementId>();
                                    placePointIds = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);

                                    for (int ptNum = 0; ptNum < obj.AdaptivePoints.Count; ptNum++)
                                    {
                                        try
                                        {
                                            ReferencePoint rp = doc.GetElement(placePointIds[ptNum]) as ReferencePoint;
                                            XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].Z, lengthDUT));
                                            if (rp != null)
                                            {
                                                XYZ vector = pt.Subtract(rp.Position);
                                                ElementTransformUtils.MoveElement(doc, rp.Id, vector);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                    }

                                    // Assign the parameters
                                    SetParameters(fi, obj.Parameters, doc);

                                    // Assign the GH InstanceGuid
                                    AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                }

                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                        t.Commit();
                    }
                }
            }
            #endregion

            #region Curve based
            else if (ro.Curves != null && ro.Curves.Count > 0)
            {

                // Find the FamilySymbol
                FamilySymbol symbol = null;
                WallType wallType = null;
                FloorType floorType = null;
                RoofType roofType = null;
                GridType gridType = null;
                bool typeFound = false;

                FilteredElementCollector famCollector = new FilteredElementCollector(doc);

                if (ro.CategoryId == -2000011)
                {
                    famCollector.OfClass(typeof(WallType));
                    foreach (WallType wt in famCollector)
                    {
                        if (wt.Name == ro.TypeName)
                        {
                            wallType = wt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000032)
                {
                    famCollector.OfClass(typeof(FloorType));
                    foreach (FloorType ft in famCollector)
                    {
                        if (ft.Name == ro.TypeName)
                        {
                            floorType = ft;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000035)
                {
                    famCollector.OfClass(typeof(RoofType));
                    foreach (RoofType rt in famCollector)
                    {
                        if (rt.Name == ro.TypeName)
                        {
                            roofType = rt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000220)
                {
                    famCollector.OfClass(typeof(GridType));
                    foreach (GridType gt in famCollector)
                    {
                        if (gt.Name == ro.TypeName)
                        {
                            gridType = gt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);
                    if (symbol != null)
                        typeFound = true;
                }



                if (typeFound)
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
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");

                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");

                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }
                            FamilyInstance fi = null;
                            try
                            {

                                foreach (RevitObject obj in revitObjects)
                                {

                                    #region single line based family
                                    if (obj.Curves.Count == 1 && obj.Curves[0].CurveType != "Circle")
                                    {

                                        LyrebirdCurve lbc = obj.Curves[0];
                                        List<LyrebirdPoint> curvePoints = lbc.ControlPoints.OrderBy(p => p.Z).ToList();
                                        // linear
                                        // can be a wall or line based family.
                                        if (obj.CategoryId == -2000011)
                                        {

                                            // draw a wall
                                            Curve crv = null;
                                            if (lbc.CurveType == "Line")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                crv = Line.CreateBound(pt1, pt2);
                                            }
                                            else if (lbc.CurveType == "Arc")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                crv = Arc.Create(pt1, pt3, pt2);
                                            }

                                            if (crv != null)
                                            {

                                                // Find the level
                                                Level lvl = GetLevel(lbc.ControlPoints, doc);

                                                double offset = 0;
                                                if (Math.Abs(UnitUtils.ConvertToInternalUnits(curvePoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                                {
                                                    offset = UnitUtils.ConvertToInternalUnits(curvePoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                                }

                                                // Create the wall
                                                Wall w = null;
                                                try
                                                {
                                                    w = Wall.Create(doc, crv, wallType.Id, lvl.Id, 10.0, offset, false, false);
                                                }
                                                catch (Exception ex)
                                                {
                                                    TaskDialog.Show("ERROR", ex.Message);
                                                }

                                                // Assign the parameters
                                                SetParameters(w, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(w, uniqueId, instanceSchema, 0, nickName);
                                            }
                                        }
                                        // See if it's a structural column
                                        else if (obj.CategoryId == -2001330)
                                        {
                                            if (symbol != null && lbc.CurveType == "Line")
                                            {
                                                XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                Curve crv = Line.CreateBound(origin, pt2);

                                                // Find the level
                                                Level lvl = GetLevel(lbc.ControlPoints, doc);

                                                // Create the column
                                                fi = doc.Create.NewFamilyInstance(origin, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.Column);

                                                // Change it to a slanted column
                                                Parameter slantParam = fi.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM);

                                                // SlantedOrVerticalColumnType has 3 options, CT_Vertical (0), CT_Angle (1), or CT_EndPoint (2)
                                                // CT_EndPoint is what we want for a line based column.
                                                slantParam.Set(2);

                                                // Set the location curve of the column to the line
                                                LocationCurve lc = fi.Location as LocationCurve;
                                                if (lc != null)
                                                {
                                                    lc.Curve = crv;
                                                }

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                            }
                                        }
                                        else if (obj.CategoryId == -2000220)
                                        {
                                            // draw a grid
                                            Grid g = null;
                                            if (lbc.CurveType == "Line")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                Line line = Line.CreateBound(pt1, pt2);
                                                try
                                                {
                                                    g = doc.Create.NewGrid(line);
                                                }
                                                catch { }
                                            }
                                            else if (lbc.CurveType == "Arc")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                Arc arc = Arc.Create(pt1, pt3, pt2);
                                                try
                                                {
                                                    g = doc.Create.NewGrid(arc);
                                                }
                                                catch { }
                                            }

                                            if (g != null)
                                            {
                                                g.ExtendToAllLevels(); ;
                                                // Assign the parameters
                                                SetParameters(g, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(g, uniqueId, instanceSchema, 0, nickName);
                                            }
                                        }
                                        else if (obj.CategoryId == -2000051)
                                        {
                                            // Draw a line
                                            Curve crv = null;
                                            Curve crv2 = null;

                                            if (lbc.CurveType == "Line")
                                            {
                                                XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                crv = Line.CreateBound(origin, pt2);
                                            }
                                            else if (lbc.CurveType == "Arc")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                crv = Arc.Create(pt1, pt3, pt2);
                                            }
                                            else if (lbc.CurveType == "Circle")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                XYZ pt4 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Z, lengthDUT));
                                                XYZ pt5 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Z, lengthDUT));
                                                Arc arc1 = Arc.Create(pt1, pt3, pt2);
                                                Arc arc2 = Arc.Create(pt3, pt5, pt4);
                                                crv = arc1;
                                                crv2 = arc2;
                                            }
                                            else if (lbc.CurveType == "Spline")
                                            {
                                                List<XYZ> controlPoints = new List<XYZ>();
                                                List<double> weights = lbc.Weights;
                                                List<double> knots = lbc.Knots;


                                                foreach (LyrebirdPoint lp in lbc.ControlPoints)
                                                {
                                                    XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(lp.X, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Z, lengthDUT));
                                                    controlPoints.Add(pt);
                                                }

                                                NurbSpline spline;
                                                if (lbc.Degree < 3)
                                                    spline = NurbSpline.Create(controlPoints, weights);
                                                else
                                                    spline = NurbSpline.Create(controlPoints, weights, knots, lbc.Degree, false, true);

                                                crv = spline;
                                            }

                                            // Check for model or detail
                                            if (obj.FamilyName == "Model Lines")
                                            {
                                                // We need a plane
                                                
                                            }
                                            else if (obj.FamilyName == "Detail Lines")
                                            {
                                                // we need the active view.
                                            }
                                        }

                                        // Otherwise create a family it using the line
                                        else
                                        {
                                            if (symbol != null)
                                            {
                                                Curve crv = null;

                                                if (lbc.CurveType == "Line")
                                                {
                                                    XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                    crv = Line.CreateBound(origin, pt2);
                                                }
                                                else if (lbc.CurveType == "Arc")
                                                {
                                                    XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                    XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                    crv = Arc.Create(pt1, pt3, pt2);
                                                }
                                                // Find the level
                                                Level lvl = GetLevel(lbc.ControlPoints, doc);

                                                // Create the family
                                                if (symbol.Category.Id.IntegerValue == -2002000)
                                                {
                                                    try
                                                    {
                                                        Line line = crv as Line;
                                                        fi = doc.Create.NewFamilyInstance(line, symbol, doc.ActiveView);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                else if (symbol.Category.Id.IntegerValue == -2001320)
                                                {
                                                    try
                                                    {
                                                        if (lbc.CurveType == "Arc")
                                                        {
                                                            XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                            fi = doc.Create.NewFamilyInstance(origin, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                            // Set the location curve of the column to the line
                                                            LocationCurve lc = fi.Location as LocationCurve;
                                                            if (lc != null)
                                                            {
                                                                lc.Curve = crv;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            fi = doc.Create.NewFamilyInstance(crv, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                else
                                                {
                                                    try
                                                    {
                                                        fi = doc.Create.NewFamilyInstance(crv, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Closed Curve Family
                                    else
                                    {
                                        // A list of curves.  These should equate a closed planar curve from GH.

                                        //TODO: For each profile type, determine if the offset is working correctly or inverted

                                        // Then determine category and create based on that.
                                        if (obj.CategoryId == -2000011)
                                        {
                                            // Create line based wall
                                            // Find the level
                                            double offset = 0;
                                            List<LyrebirdPoint> allPoints = new List<LyrebirdPoint>();
                                            foreach (LyrebirdCurve lc in obj.Curves)
                                            {
                                                foreach (LyrebirdPoint lp in lc.ControlPoints)
                                                {
                                                    allPoints.Add(lp);
                                                }
                                            }
                                            allPoints.Sort((x, y) => x.Z.CompareTo(y.Z));

                                            Level lvl = GetLevel(allPoints, doc);

                                            if (Math.Abs(UnitUtils.ConvertToInternalUnits(allPoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = UnitUtils.ConvertToInternalUnits(allPoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            List<Curve> crvArray = new List<Curve>();
                                            try
                                            {
                                                foreach (LyrebirdCurve lbc in obj.Curves)
                                                {
                                                    if (lbc.CurveType == "Circle")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                        XYZ pt4 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Z, lengthDUT));
                                                        XYZ pt5 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Z, lengthDUT));
                                                        Arc arc1 = Arc.Create(pt1, pt3, pt2);
                                                        Arc arc2 = Arc.Create(pt3, pt5, pt4);
                                                        crvArray.Add(arc1);
                                                        crvArray.Add(arc2);
                                                    }
                                                    else if (lbc.CurveType == "Arc")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                        Arc arc = Arc.Create(pt1, pt3, pt2);
                                                        crvArray.Add(arc);
                                                    }
                                                    else if (lbc.CurveType == "Line")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        Line line = Line.CreateBound(pt1, pt2);
                                                        crvArray.Add(line);
                                                    }
                                                    else if (lbc.CurveType == "Spline")
                                                    {
                                                        List<XYZ> controlPoints = new List<XYZ>();
                                                        List<double> weights = lbc.Weights;
                                                        List<double> knots = lbc.Knots;


                                                        foreach (LyrebirdPoint lp in lbc.ControlPoints)
                                                        {
                                                            XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(lp.X, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Z, lengthDUT));
                                                            controlPoints.Add(pt);
                                                        }

                                                        NurbSpline spline;
                                                        if (lbc.Degree < 3)
                                                            spline = NurbSpline.Create(controlPoints, weights);
                                                        else
                                                            spline = NurbSpline.Create(controlPoints, weights, knots, lbc.Degree, false, true);

                                                        crvArray.Add(spline);
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                TaskDialog.Show("ERROR", ex.Message);
                                            }

                                            // Create the floor
                                            Wall w = Wall.Create(doc, crvArray, wallType.Id, lvl.Id, false);
                                            if (Math.Abs(offset - 0) > double.Epsilon)
                                            {
                                                Parameter p = w.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
                                                p.Set(offset);
                                            }

                                            // Assign the parameters
                                            SetParameters(w, obj.Parameters, doc);

                                            // Assign the GH InstanceGuid
                                            AssignGuid(w, uniqueId, instanceSchema, 0, nickName);

                                        }
                                        else if (obj.CategoryId == -2000032)
                                        {
                                            // Create a profile based floor
                                            // Find the level
                                            Level lvl = GetLevel(obj.Curves[0].ControlPoints, doc);

                                            double offset = 0;
                                            if (Math.Abs(UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            CurveArray crvArray;
                                            Floor flr;
                                            List<Opening> flrOpenings = new List<Opening>();
                                            if (obj.CurveIds != null && obj.CurveIds.Count > 0)
                                            {
                                                // get the main profile
                                                int crvCount = obj.CurveIds[0];
                                                List<LyrebirdCurve> primaryCurves = obj.Curves.GetRange(0, crvCount);
                                                crvArray = GetCurveArray(primaryCurves);
                                                flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);

                                                // TODO: You cannot create holes in an element with the API without using openings rather than interior closed curves.
                                                // Evaluate with later versions if it's worth creating the openings and updating them
                                                // Create the openings associated with it.
                                                //int start = crvCount - 1;
                                                //for (int i = 1; i < obj.CurveIds.Count; i++)
                                                //{
                                                //    List<LyrebirdCurve> interiorCurves = obj.Curves.GetRange(start, obj.CurveIds[i]);
                                                //    start += interiorCurves.Count;
                                                //    CurveArray openingArray = GetCurveArray(interiorCurves);
                                                //    try
                                                //    {
                                                //        SubTransaction st2 = new SubTransaction(doc);
                                                //        st2.Start();
                                                //        Opening opening = doc.Create.NewOpening(flr, openingArray, false);
                                                //        st2.Commit();
                                                //        flrOpenings.Add(opening);

                                                //    }
                                                //    catch (Exception ex)
                                                //    {
                                                //        TaskDialog.Show("TEST", ex.Message);
                                                //    }
                                                //}
                                            }
                                            else
                                            {
                                                crvArray = GetCurveArray(obj.Curves);
                                                flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);
                                            }

                                            // Create the floor
                                            //flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);

                                            if (Math.Abs(offset - 0) > double.Epsilon)
                                            {
                                                Parameter p = flr.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                                                p.Set(offset);
                                            }

                                            // Assign the parameters
                                            SetParameters(flr, obj.Parameters, doc);

                                            // Assign the GH InstanceGuid
                                            AssignGuid(flr, uniqueId, instanceSchema, 0, nickName);

                                        }
                                        else if (obj.CategoryId == -2000035)
                                        {
                                            // Create a RoofExtrusion
                                            // Find the level
                                            Level lvl = GetLevel(obj.Curves[0].ControlPoints, doc);

                                            double offset = 0;
                                            if (Math.Abs(UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            CurveArray crvArray = GetCurveArray(obj.Curves);

                                            // Create the roof
                                            FootPrintRoof roof = null;
                                            ModelCurveArray roofProfile = new ModelCurveArray();
                                            try
                                            {
                                                roof = doc.Create.NewFootPrintRoof(crvArray, lvl, roofType, out roofProfile);
                                            }
                                            catch (Exception ex)
                                            {
                                                TaskDialog.Show("ERROR", ex.Message);
                                            }
                                            if (Math.Abs(offset - 0) > double.Epsilon)
                                            {
                                                Parameter p = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                                                p.Set(offset);
                                            }

                                            // Assign the parameters
                                            SetParameters(roof, obj.Parameters, doc);

                                            // Assign the GH InstanceGuid
                                            AssignGuid(roof, uniqueId, instanceSchema, 0, nickName);
                                        }
                                    }
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Error", ex.ToString());
                                Debug.WriteLine(ex.Message);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                        t.Commit();
                    }
                }


            }
            #endregion
        }

        private void ModifyObjects(List<RevitObject> existingObjects, List<ElementId> existingElems, Document doc, Guid uniqueId, bool profileWarning, string nickName, int runId)
        {
            // Create new Revit objects.
            //List<LyrebirdId> newUniqueIds = new List<LyrebirdId>();

            // Determine what kind of object we're creating.
            RevitObject ro = existingObjects[0];


            #region Normal Origin based FamilyInstance
            // Modify origin based family instances
            if (ro.Origin != null)
            {
                // Find the FamilySymbol
                FamilySymbol symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);

                LevelType lt = FindLevelType(ro.TypeName, doc);

                GridType gt = FindGridType(ro.TypeName, doc);

                if (symbol != null || lt != null)
                {
                    // Get the hosting ID from the family.
                    Family fam = null;
                    Parameter hostParam = null;
                    int hostBehavior = 0;

                    try
                    {
                        fam = symbol.Family;
                        hostParam = fam.get_Parameter(BuiltInParameter.FAMILY_HOSTING_BEHAVIOR);
                        hostBehavior = hostParam.AsInteger();
                    }
                    catch{}

                    //FamilyInstance existingInstance = doc.GetElement(existingElems[0]) as FamilyInstance;
                    
                    using (Transaction t = new Transaction(doc, "Lyrebird Modify Objects"))
                    {
                        t.Start();
                        try
                        {
                            Schema instanceSchema = null;
                            try
                            {
                                instanceSchema = Schema.Lookup(instanceSchemaGUID);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");
                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");
                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }

                            FamilyInstance fi = null;
                            XYZ origin = XYZ.Zero;

                            if (lt != null)
                            {
                                for (int i = 0; i < existingObjects.Count; i++)
                                {
                                    RevitObject obj = existingObjects[i];
                                    Level lvl = doc.GetElement(existingElems[i]) as Level;

                                    if (lvl.ProjectElevation != (UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT)))
                                    {
                                        double offset = lvl.ProjectElevation - lvl.ProjectElevation;
                                        lvl.Elevation = (UnitUtils.ConvertToInternalUnits(obj.Origin.Z + offset, lengthDUT));
                                    }

                                    SetParameters(lvl, obj.Parameters, doc);
                                }
                            }
                            else if (hostBehavior == 0)
                            {

                                for (int i = 0; i < existingObjects.Count; i++)
                                {
                                    RevitObject obj = existingObjects[i];
                                    fi = doc.GetElement(existingElems[i]) as FamilyInstance;

                                    // Change the family and symbol if necessary
                                    if (fi != null && (fi.Symbol.Family.Name != symbol.Family.Name || fi.Symbol.Name != symbol.Name))
                                    {
                                        try
                                        {
                                            fi.Symbol = symbol;
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                    }

                                    try
                                    {
                                        // Move family
                                        origin = new XYZ(UnitUtils.ConvertToInternalUnits(obj.Origin.X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT));
                                        if (fi != null)
                                        {
                                            LocationPoint lp = fi.Location as LocationPoint;
                                            if (lp != null)
                                            {
                                                XYZ oldLoc = lp.Point;
                                                XYZ translation = origin.Subtract(oldLoc);
                                                ElementTransformUtils.MoveElement(doc, fi.Id, translation);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Error", ex.Message);
                                    }

                                    // Rotate
                                    if (obj.Orientation != null)
                                    {
                                        if (Math.Round(Math.Abs(obj.Orientation.Z - 0), 10) < double.Epsilon)
                                        {
                                            XYZ orientation = fi.FacingOrientation;
                                            orientation = orientation.Multiply(-1);
                                            XYZ incomingOrientation = new XYZ(obj.Orientation.X, obj.Orientation.Y, obj.Orientation.Z);
                                            XYZ normalVector = new XYZ(0, -1, 0);

                                            double currentAngle = 0;
                                            if (orientation.X < 0 && orientation.Y < 0)
                                            {
                                                currentAngle = (2 * Math.PI) - normalVector.AngleTo(orientation);
                                            }
                                            else if (orientation.Y == 0 && orientation.X < 0)
                                            {
                                                currentAngle = 1.5 * Math.PI;
                                            }
                                            else if (orientation.X < 0)
                                            {
                                                currentAngle = (Math.PI - normalVector.AngleTo(orientation)) + Math.PI;
                                            }
                                            else
                                            {
                                                currentAngle = normalVector.AngleTo(orientation);
                                            }

                                            double incomingAngle = 0;
                                            if (incomingOrientation.X < 0 && incomingOrientation.Y < 0)
                                            {
                                                incomingAngle = (2 * Math.PI) - normalVector.AngleTo(incomingOrientation);
                                            }
                                            else if (incomingOrientation.Y == 0 && incomingOrientation.X < 0)
                                            {
                                                incomingAngle = 1.5 * Math.PI;
                                            }
                                            else if (incomingOrientation.X < 0)
                                            {
                                                incomingAngle = (Math.PI - normalVector.AngleTo(incomingOrientation)) + Math.PI;
                                            }
                                            else
                                            {
                                                incomingAngle = normalVector.AngleTo(incomingOrientation);
                                            }
                                            double angle = incomingAngle - currentAngle;
                                            //TaskDialog.Show("Test", "CurrentAngle: " + currentAngle.ToString() + "\nIncoming Angle: " + incomingAngle.ToString() + "\nResulting Rotation: " + angle.ToString() +
                                            //    "\nFacingOrientation: " + orientation.ToString() + "\nIncoming Orientation: " + incomingOrientation.ToString());
                                            Line axis = Line.CreateBound(origin, origin + XYZ.BasisZ);
                                            ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);
                                        }
                                    }

                                    SetParameters(fi, obj.Parameters, doc);
                                }
                            }
                            else
                            {
                                for (int i = 0; i < existingObjects.Count; i++)
                                {
                                    RevitObject obj = existingObjects[i];
                                    fi = doc.GetElement(existingElems[i]) as FamilyInstance;

                                    // Change the family and symbol if necessary
                                    if (fi.Symbol.Family.Name != symbol.Family.Name || fi.Symbol.Name != symbol.Name)
                                    {
                                        try
                                        {
                                            fi.Symbol = symbol;
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                    }

                                    origin = new XYZ(UnitUtils.ConvertToInternalUnits(obj.Origin.X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.Origin.Z, lengthDUT));

                                    // Find the level
                                    List<LyrebirdPoint> lbPoints = new List<LyrebirdPoint> { obj.Origin };
                                    Level lvl = GetLevel(lbPoints, doc);

                                    // Get the host
                                    if (hostBehavior == 5)
                                    {
                                        // Face based family.  Find the face and create the element
                                        XYZ normVector = new XYZ(obj.Orientation.X, obj.Orientation.Y, obj.Orientation.Z);
                                        XYZ faceVector;
                                        if (obj.FaceOrientation != null)
                                        {
                                            faceVector = new XYZ(obj.FaceOrientation.X, obj.FaceOrientation.Y, obj.FaceOrientation.Z);
                                        }
                                        else
                                        {
                                            faceVector = XYZ.BasisZ;
                                        }
                                        Face face = FindFace(origin, normVector, doc);
                                        if (face != null)
                                        {
                                            if (face.Reference.ElementId == fi.HostFace.ElementId)
                                            {
                                                //fi = doc.Create.NewFamilyInstance(face, origin, faceVector, symbol);
                                                // Just move the host and update the parameters as needed.
                                                LocationPoint lp = fi.Location as LocationPoint;
                                                if (lp != null)
                                                {
                                                    XYZ oldLoc = lp.Point;
                                                    XYZ translation = origin.Subtract(oldLoc);
                                                    ElementTransformUtils.MoveElement(doc, fi.Id, translation);
                                                }

                                                SetParameters(fi, obj.Parameters, doc);
                                            }
                                            else
                                            {
                                                FamilyInstance origInst = fi;
                                                fi = doc.Create.NewFamilyInstance(face, origin, faceVector, symbol);

                                                foreach (Parameter p in origInst.Parameters)
                                                {
                                                    try
                                                    {
                                                        Parameter newParam = fi.get_Parameter(p.GUID);
                                                        if (newParam != null)
                                                        {
                                                            switch (newParam.StorageType)
                                                            {
                                                                case StorageType.Double:
                                                                    newParam.Set(p.AsDouble());
                                                                    break;
                                                                case StorageType.ElementId:
                                                                    newParam.Set(p.AsElementId());
                                                                    break;
                                                                case StorageType.Integer:
                                                                    newParam.Set(p.AsInteger());
                                                                    break;
                                                                case StorageType.String:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                                default:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                // Delete the original instance of the family
                                                doc.Delete(origInst.Id);

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // typical hosted family.  Can be wall, floor, roof or ceiling.
                                        ElementId host = FindHost(origin, hostBehavior, doc);
                                        if (host != null)
                                        {
                                            if (host.IntegerValue != fi.Host.Id.IntegerValue)
                                            {
                                                // We'll have to recreate the element
                                                FamilyInstance origInst = fi;
                                                fi = doc.Create.NewFamilyInstance(origin, symbol, doc.GetElement(host), lvl, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                foreach (Parameter p in origInst.Parameters)
                                                {
                                                    try
                                                    {
                                                        Parameter newParam = fi.LookupParameter(p.Definition.Name);
                                                        if (newParam != null)
                                                        {
                                                            switch (newParam.StorageType)
                                                            {
                                                                case StorageType.Double:
                                                                    newParam.Set(p.AsDouble());
                                                                    break;
                                                                case StorageType.ElementId:
                                                                    newParam.Set(p.AsElementId());
                                                                    break;
                                                                case StorageType.Integer:
                                                                    newParam.Set(p.AsInteger());
                                                                    break;
                                                                case StorageType.String:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                                default:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                // Delete the original instance of the family
                                                doc.Delete(origInst.Id);

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(fi, uniqueId, instanceSchema, runId, nickName);
                                            }

                                            else
                                            {
                                                // Just move the host and update the parameters as needed.
                                                LocationPoint lp = fi.Location as LocationPoint;
                                                if (lp != null)
                                                {
                                                    XYZ oldLoc = lp.Point;
                                                    XYZ translation = origin.Subtract(oldLoc);
                                                    ElementTransformUtils.MoveElement(doc, fi.Id, translation);
                                                }

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);
                                            }
                                        }
                                    }

                                    // Assign the parameters
                                    SetParameters(fi, obj.Parameters, doc);
                                }
                                // delete the host finder
                                ElementId hostFinderFamily = hostFinder.Symbol.Family.Id;
                                doc.Delete(hostFinder.Id);
                                doc.Delete(hostFinderFamily);
                                hostFinder = null;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }

                        t.Commit();
                    }
                }
            }

            #endregion


            #region Adaptive Components
            FamilyInstance adaptInst;
            try
            {
                adaptInst = doc.GetElement(existingElems[0]) as FamilyInstance;
            }
            catch
            {
                adaptInst = null;
            }
            if (adaptInst != null && AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(adaptInst))
            {
                // Find the FamilySymbol
                FamilySymbol symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);

                if (symbol != null)
                {
                    using (Transaction t = new Transaction(doc, "Lyrebird Modify Objects"))
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
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");
                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");
                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }

                            try
                            {
                                for (int i = 0; i < existingElems.Count; i++)
                                {
                                    RevitObject obj = existingObjects[i];

                                    FamilyInstance fi = doc.GetElement(existingElems[i]) as FamilyInstance;

                                    // Change the family and symbol if necessary
                                    if (fi.Symbol.Family.Name != symbol.Family.Name || fi.Symbol.Name != symbol.Name)
                                    {
                                        try
                                        {
                                            fi.Symbol = symbol;
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                    }

                                    IList<ElementId> placePointIds = new List<ElementId>();
                                    placePointIds = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);

                                    for (int ptNum = 0; ptNum < obj.AdaptivePoints.Count; ptNum++)
                                    {
                                        try
                                        {
                                            ReferencePoint rp = doc.GetElement(placePointIds[ptNum]) as ReferencePoint;
                                            XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].X, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(obj.AdaptivePoints[ptNum].Z, lengthDUT));
                                            XYZ vector = pt.Subtract(rp.Position);
                                            ElementTransformUtils.MoveElement(doc, rp.Id, vector);
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                    }

                                    // Assign the parameters
                                    SetParameters(fi, obj.Parameters, doc);
                                }

                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                        t.Commit();
                    }
                }

            }

            #endregion


            #region Curve based components
            if (ro.Curves != null && ro.Curves.Count > 0)
            {
                // Find the FamilySymbol
                FamilySymbol symbol = null;
                WallType wallType = null;
                FloorType floorType = null;
                RoofType roofType = null;
                GridType gridType = null;
                bool typeFound = false;

                FilteredElementCollector famCollector = new FilteredElementCollector(doc);

                if (ro.CategoryId == -2000011)
                {
                    famCollector.OfClass(typeof(WallType));
                    foreach (WallType wt in famCollector)
                    {
                        if (wt.Name == ro.TypeName)
                        {
                            wallType = wt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000032)
                {
                    famCollector.OfClass(typeof(FloorType));
                    foreach (FloorType ft in famCollector)
                    {
                        if (ft.Name == ro.TypeName)
                        {
                            floorType = ft;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000035)
                {
                    famCollector.OfClass(typeof(RoofType));
                    foreach (RoofType rt in famCollector)
                    {
                        if (rt.Name == ro.TypeName)
                        {
                            roofType = rt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else if (ro.CategoryId == -2000220)
                {
                    famCollector.OfClass(typeof(GridType));
                    foreach (GridType gt in famCollector)
                    {
                        if (gt.Name == ro.TypeName)
                        {
                            gridType = gt;
                            typeFound = true;
                            break;
                        }
                    }
                }
                else
                {
                    symbol = FindFamilySymbol(ro.FamilyName, ro.TypeName, doc);
                    if (symbol != null)
                        typeFound = true;
                }



                if (typeFound)
                {
                    
                    using (Transaction t = new Transaction(doc, "Lyrebird Modify Objects"))
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
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                            if (instanceSchema == null)
                            {
                                SchemaBuilder sb = new SchemaBuilder(instanceSchemaGUID);
                                sb.SetWriteAccessLevel(AccessLevel.Vendor);
                                sb.SetReadAccessLevel(AccessLevel.Public);
                                sb.SetVendorId("LMNA");

                                // Create the field to store the data in the family
                                FieldBuilder guidFB = sb.AddSimpleField("InstanceID", typeof(string));
                                guidFB.SetDocumentation("Component instance GUID from Grasshopper");
                                // Create a filed to store the run number
                                FieldBuilder runIDFB = sb.AddSimpleField("RunID", typeof(int));
                                runIDFB.SetDocumentation("RunID for when multiple runs are created from the same data");
                                // Create a field to store the GH component nickname.
                                FieldBuilder nickNameFB = sb.AddSimpleField("NickName", typeof(string));
                                nickNameFB.SetDocumentation("Component NickName from Grasshopper");

                                sb.SetSchemaName("LMNAInstanceGUID");
                                instanceSchema = sb.Finish();
                            }
                            FamilyInstance fi = null;
                            Grid grid = null;
                            
                            try
                            {
                                bool supress = Properties.Settings.Default.suppressWarning;
                                bool supressedReplace = false;
                                bool supressedModify = true;
                                for (int i = 0; i < existingObjects.Count; i++)
                                {
                                    RevitObject obj = existingObjects[i];
                                    if (obj.CategoryId != -2000011 && obj.CategoryId != -2000032 && obj.CategoryId != -2000035 && obj.CategoryId != -2000220)
                                    {
                                        fi = doc.GetElement(existingElems[i]) as FamilyInstance;

                                        // Change the family and symbol if necessary
                                        if (fi.Symbol.Family.Name != symbol.Family.Name || fi.Symbol.Name != symbol.Name)
                                        {
                                            try
                                            {
                                                fi.Symbol = symbol;
                                            }
                                            catch (Exception ex)
                                            {
                                                Debug.WriteLine(ex.Message);
                                            }
                                        }
                                    }
                                    else if (obj.CategoryId == -2000220)
                                    {
                                        grid = doc.GetElement(existingElems[i]) as Grid;

                                        // Get the grid location and compare against the incoming curve
                                        Curve gridCrv = grid.Curve;
                                        LyrebirdCurve lbc = obj.Curves[0];
                                        try
                                        {
                                            Arc arc = gridCrv as Arc;
                                            if (arc != null && lbc.CurveType == "Arc")
                                            {
                                                // Test that the arcs are similar
                                                XYZ startPoint = arc.GetEndPoint(0);
                                                XYZ endPoint = arc.GetEndPoint(1);
                                                XYZ centerPoint = arc.Center;
                                                double rad = arc.Radius;

                                                XYZ lbcStartPt = new XYZ(lbc.ControlPoints[0].X, lbc.ControlPoints[0].Y, lbc.ControlPoints[0].Z);
                                                XYZ lbcEndPt = new XYZ(lbc.ControlPoints[1].X, lbc.ControlPoints[1].Y, lbc.ControlPoints[1].Z);
                                                XYZ lbcMidPt = new XYZ(lbc.ControlPoints[2].X, lbc.ControlPoints[2].Y, lbc.ControlPoints[2].Z);
                                                Arc lbcArc = Arc.Create(lbcStartPt, lbcEndPt, lbcMidPt);
                                                XYZ lbcCenterPt = lbcArc.Center;
                                                double lbcRad = lbcArc.Radius;

                                                if (centerPoint.DistanceTo(lbcCenterPt) < 0.001 && lbcRad == rad && startPoint.DistanceTo(lbcStartPt) < 0.001)
                                                {
                                                    // Do not create
                                                }
                                                else
                                                {
                                                    // Delete the grid and rebuild it with the new curve.
                                                }
                                            }
                                            else
                                            {
                                                // Probably need to rebuild the curve
                                            }
                                        }
                                        catch { }


                                        try
                                        {
                                            Line line = gridCrv as Line;
                                            if (line != null && lbc.CurveType == "Line")
                                            {
                                                // Test that the arcs are similar
                                                XYZ startPoint = line.GetEndPoint(0);
                                                XYZ endPoint = line.GetEndPoint(1);

                                                XYZ lbcStartPt = new XYZ(lbc.ControlPoints[0].X, lbc.ControlPoints[0].Y, lbc.ControlPoints[0].Z);
                                                XYZ lbcEndPt = new XYZ(lbc.ControlPoints[1].X, lbc.ControlPoints[1].Y, lbc.ControlPoints[1].Z);

                                                if (endPoint.DistanceTo(lbcEndPt) < 0.001 && startPoint.DistanceTo(lbcStartPt) < 0.001)
                                                {
                                                    // Do not create
                                                }
                                                else
                                                {
                                                    // Delete the grid and rebuild it with the new curve.
                                                }
                                            }
                                            else
                                            {
                                                // Probably need to rebuild the curve
                                            }
                                        }
                                        catch { }

                                        if (grid.GridType.Name != gridType.Name)
                                        {
                                            try
                                            {
                                                grid.GridType = gridType;
                                            }
                                            catch (Exception ex)
                                            {
                                                TaskDialog.Show("Error", ex.Message);
                                                Debug.WriteLine(ex.Message);
                                            }
                                        }
                                    }
                                    #region single line based family
                                    if (obj.Curves.Count == 1 && obj.Curves[0].CurveType != "Circle")
                                    {

                                        LyrebirdCurve lbc = obj.Curves[0];
                                        List<LyrebirdPoint> curvePoints = lbc.ControlPoints.OrderBy(p => p.Z).ToList();
                                        // linear
                                        // can be a wall or line based family.

                                        // Wall objects
                                        if (obj.CategoryId == -2000011)
                                        {
                                            // draw a wall
                                            Curve crv = null;
                                            if (lbc.CurveType == "Line")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                crv = Line.CreateBound(pt1, pt2);
                                            }
                                            else if (lbc.CurveType == "Arc")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                crv = Arc.Create(pt1, pt3, pt2);
                                            }

                                            if (crv != null)
                                            {

                                                // Find the level
                                                Level lvl = GetLevel(lbc.ControlPoints, doc);

                                                double offset = 0;
                                                if (Math.Abs(UnitUtils.ConvertToInternalUnits(curvePoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                                {
                                                    offset = lvl.ProjectElevation - UnitUtils.ConvertToInternalUnits(curvePoints[0].Z, lengthDUT);
                                                }

                                                // Modify the wall
                                                Wall w = null;
                                                try
                                                {
                                                    w = doc.GetElement(existingElems[i]) as Wall;
                                                    LocationCurve lc = w.Location as LocationCurve;
                                                    lc.Curve = crv;

                                                    // Change the family and symbol if necessary
                                                    if (w.WallType.Name != wallType.Name)
                                                    {
                                                        try
                                                        {
                                                            w.WallType = wallType;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Debug.WriteLine(ex.Message);
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    TaskDialog.Show("ERROR", ex.Message);
                                                }


                                                // Assign the parameters
                                                SetParameters(w, obj.Parameters, doc);
                                            }
                                        }
                                        // See if it's a structural column
                                        else if (obj.CategoryId == -2001330)
                                        {
                                            if (symbol != null && lbc.CurveType == "Line")
                                            {
                                                Curve crv = null;
                                                XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                crv = Line.CreateBound(origin, pt2);

                                                // Find the level
                                                //Level lvl = GetLevel(lbc.ControlPoints, doc);

                                                // Create the column
                                                //fi = doc.Create.NewFamilyInstance(origin, symbol, lvl, Autodesk.Revit.DB.Structure.StructuralType.Column);

                                                // Change it to a slanted column
                                                Parameter slantParam = fi.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM);

                                                // SlantedOrVerticalColumnType has 3 options, CT_Vertical (0), CT_Angle (1), or CT_EndPoint (2)
                                                // CT_EndPoint is what we want for a line based column.
                                                slantParam.Set(2);

                                                // Set the location curve of the column to the line
                                                LocationCurve lc = fi.Location as LocationCurve;
                                                if (lc != null)
                                                {
                                                    lc.Curve = crv;
                                                }

                                                // Change the family and symbol if necessary
                                                if (fi.Symbol.Name != symbol.Name)
                                                {
                                                    try
                                                    {
                                                        fi.Symbol = symbol;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }

                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);
                                            }
                                        }

                                        else if (obj.CategoryId == -2000220)
                                        {
                                            // draw a grid
                                            Curve crv = null;
                                            if (lbc.CurveType == "Line")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                crv = Line.CreateBound(pt1, pt2);
                                            }
                                            else if (lbc.CurveType == "Arc")
                                            {
                                                XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                crv = Arc.Create(pt1, pt3, pt2);
                                            }

                                            if (crv != null && grid != null)
                                            {
                                                // Determine if it's possible to edit the grid curve or if it needs to be deleted/replaced.

                                                // Assign the parameters
                                                SetParameters(grid, obj.Parameters, doc);
                                            }
                                        }

                                        // Otherwise create a family it using the line
                                        else
                                        {
                                            if (symbol != null)
                                            {
                                                Curve crv = null;

                                                if (lbc.CurveType == "Line")
                                                {
                                                    XYZ origin = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                    crv = Line.CreateBound(origin, pt2);
                                                }
                                                else if (lbc.CurveType == "Arc")
                                                {
                                                    XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                    XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                    crv = Arc.Create(pt1, pt3, pt2);
                                                }

                                                // Change the family and symbol if necessary
                                                if (fi.Symbol.Name != symbol.Name)
                                                {
                                                    try
                                                    {
                                                        fi.Symbol = symbol;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }

                                                try
                                                {
                                                    LocationCurve lc = fi.Location as LocationCurve;
                                                    lc.Curve = crv;
                                                }
                                                catch (Exception ex)
                                                {
                                                    Debug.WriteLine(ex.Message);
                                                }
                                                // Assign the parameters
                                                SetParameters(fi, obj.Parameters, doc);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Closed Curve Family
                                    else
                                    {
                                        bool replace = false;
                                        if (supress)
                                        {
                                            if (supressedReplace)
                                            {
                                                replace = true;
                                            }
                                            else
                                            {
                                                replace = false;
                                            }
                                        }
                                        if (profileWarning && !supress)
                                        {
                                            TaskDialog warningDlg = new TaskDialog("Warning")
                                            {
                                                MainInstruction = "Profile based Elements warning",
                                                MainContent =
                                                  "Elements that require updates to a profile sketch may not be updated if the number of curves in the sketch differs from the incoming curves." +
                                                  "  In such cases the element and will be deleted and replaced with new elements." +
                                                  "  Doing so will cause the loss of any elements hosted to the original instance. How would you like to proceed"
                                            };

                                            warningDlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Replace the existing elements, understanding hosted elements may be lost");
                                            warningDlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Only updated parameter information and not profile or location information");
                                            warningDlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
                                            //warningDlg.VerificationText = "Supress similar warnings";

                                            TaskDialogResult result = warningDlg.Show();
                                            if (result == TaskDialogResult.CommandLink1)
                                            {
                                                replace = true;
                                                supressedReplace = true;
                                                supressedModify = true;
                                                //supress = warningDlg.WasVerificationChecked();
                                            }
                                            if (result == TaskDialogResult.CommandLink2)
                                            {
                                                supressedReplace = false;
                                                supressedModify = true;
                                                //supress = warningDlg.WasVerificationChecked();
                                            }
                                            if (result == TaskDialogResult.CommandLink3)
                                            {
                                                supressedReplace = false;
                                                supressedModify = false;
                                                //supress = warningDlg.WasVerificationChecked();
                                            }
                                        }
                                        // A list of curves.  These should equate a closed planar curve from GH.
                                        // Determine category and create based on that.
                                        #region walls
                                        if (obj.CategoryId == -2000011)
                                        {
                                            // Create line based wall
                                            // Find the level
                                            Level lvl = null;
                                            double offset = 0;
                                            List<LyrebirdPoint> allPoints = new List<LyrebirdPoint>();
                                            foreach (LyrebirdCurve lc in obj.Curves)
                                            {
                                                foreach (LyrebirdPoint lp in lc.ControlPoints)
                                                {
                                                    allPoints.Add(lp);
                                                }
                                            }
                                            allPoints.Sort((x, y) => x.Z.CompareTo(y.Z));

                                            lvl = GetLevel(allPoints, doc);

                                            if (Math.Abs(allPoints[0].Z - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = allPoints[0].Z - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            List<Curve> crvArray = new List<Curve>();
                                            try
                                            {
                                                foreach (LyrebirdCurve lbc in obj.Curves)
                                                {
                                                    if (lbc.CurveType == "Circle")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                        XYZ pt4 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Z, lengthDUT));
                                                        XYZ pt5 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Z, lengthDUT));
                                                        Arc arc1 = Arc.Create(pt1, pt3, pt2);
                                                        Arc arc2 = Arc.Create(pt3, pt5, pt4);
                                                        crvArray.Add(arc1);
                                                        crvArray.Add(arc2);
                                                    }
                                                    else if (lbc.CurveType == "Arc")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                                                        Arc arc = Arc.Create(pt1, pt3, pt2);
                                                        crvArray.Add(arc);
                                                    }
                                                    else if (lbc.CurveType == "Line")
                                                    {
                                                        XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                                                        XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                                                        Line line = Line.CreateBound(pt1, pt2);
                                                        crvArray.Add(line);
                                                    }
                                                    else if (lbc.CurveType == "Spline")
                                                    {
                                                        List<XYZ> controlPoints = new List<XYZ>();
                                                        List<double> weights = lbc.Weights;
                                                        List<double> knots = lbc.Knots;

                                                        foreach (LyrebirdPoint lp in lbc.ControlPoints)
                                                        {
                                                            XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(lp.X, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Z, lengthDUT));
                                                            controlPoints.Add(pt);
                                                        }

                                                        NurbSpline spline;
                                                        if (lbc.Degree < 3)
                                                            spline = NurbSpline.Create(controlPoints, weights);
                                                        else
                                                            spline = NurbSpline.Create(controlPoints, weights, knots, lbc.Degree, false, true);

                                                        crvArray.Add(spline);
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                TaskDialog.Show("ERROR", ex.Message);
                                            }

                                            // Create the wall
                                            Wall w = null;
                                            if (replace)
                                            {
                                                Wall origWall = doc.GetElement(existingElems[i]) as Wall;

                                                // Find the model curves for the original wall
                                                ICollection<ElementId> ids;
                                                using (SubTransaction st = new SubTransaction(doc))
                                                {
                                                    st.Start();
                                                    ids = doc.Delete(origWall.Id);
                                                    st.RollBack();
                                                }

                                                List<ModelCurve> mLines = new List<ModelCurve>();
                                                foreach (ElementId id in ids)
                                                {
                                                    Element e = doc.GetElement(id);
                                                    if (e is ModelCurve)
                                                    {
                                                        mLines.Add(e as ModelCurve);
                                                    }
                                                }

                                                // Walls don't appear to be updatable like floors and roofs
                                                //if (mLines.Count != crvArray.Count)
                                                //{

                                                w = Wall.Create(doc, crvArray, wallType.Id, lvl.Id, false);

                                                foreach (Parameter p in origWall.Parameters)
                                                {
                                                    try
                                                    {
                                                        Parameter newParam = w.LookupParameter(p.Definition.Name);
                                                        if (newParam != null)
                                                        {
                                                            switch (newParam.StorageType)
                                                            {
                                                                case StorageType.Double:
                                                                    newParam.Set(p.AsDouble());
                                                                    break;
                                                                case StorageType.ElementId:
                                                                    newParam.Set(p.AsElementId());
                                                                    break;
                                                                case StorageType.Integer:
                                                                    newParam.Set(p.AsInteger());
                                                                    break;
                                                                case StorageType.String:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                                default:
                                                                    newParam.Set(p.AsString());
                                                                    break;
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        //TaskDialog.Show("Errorsz", ex.Message);
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }

                                                if (Math.Abs(offset - 0) > double.Epsilon)
                                                {
                                                    Parameter p = w.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
                                                    p.Set(offset);
                                                }
                                                doc.Delete(origWall.Id);

                                                // Assign the parameters
                                                SetParameters(w, obj.Parameters, doc);

                                                // Assign the GH InstanceGuid
                                                AssignGuid(w, uniqueId, instanceSchema, runId, nickName);
                                            }
                                            else if (supressedModify) // Just update the parameters and don't change the wall
                                            {
                                                w = doc.GetElement(existingElems[i]) as Wall;

                                                // Change the family and symbol if necessary
                                                if (w.WallType.Name != wallType.Name)
                                                {
                                                    try
                                                    {
                                                        w.WallType = wallType;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }

                                                // Assign the parameters
                                                SetParameters(w, obj.Parameters, doc);
                                            }
                                        }
                                        #endregion

                                        #region floors
                                        else if (obj.CategoryId == -2000032)
                                        {
                                            // Create a profile based floor
                                            // Find the level
                                            Level lvl = GetLevel(obj.Curves[0].ControlPoints, doc);

                                            double offset = 0;
                                            if (Math.Abs(UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            CurveArray crvArray = GetCurveArray(obj.Curves);

                                            // Create the floor
                                            Floor flr = null;
                                            if (replace)
                                            {
                                                Floor origFloor = doc.GetElement(existingElems[i]) as Floor;

                                                // Find the model curves for the original wall
                                                ICollection<ElementId> ids;
                                                using (SubTransaction st = new SubTransaction(doc))
                                                {
                                                    st.Start();

                                                    ids = doc.Delete(origFloor.Id);
                                                    st.RollBack();
                                                }

                                                // Get only the modelcurves
                                                List<ModelCurve> mLines = new List<ModelCurve>();
                                                foreach (ElementId id in ids)
                                                {
                                                    Element e = doc.GetElement(id);
                                                    if (e is ModelCurve)
                                                    {
                                                        mLines.Add(e as ModelCurve);
                                                    }
                                                }

                                                // Floors have an extra modelcurve for the SpanDirection.  Remove the last Item to get rid of it.
                                                mLines.RemoveAt(mLines.Count - 1);
                                                if (mLines.Count != crvArray.Size) // The sketch is different from the incoming curves so floor is recreated
                                                {
                                                    if (obj.CurveIds != null && obj.CurveIds.Count > 0)
                                                    {
                                                        // get the main profile
                                                        int crvCount = obj.CurveIds[0];
                                                        List<LyrebirdCurve> primaryCurves = obj.Curves.GetRange(0, crvCount);
                                                        crvArray = GetCurveArray(primaryCurves);
                                                        flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);
                                                    }
                                                    else
                                                    {
                                                        flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);
                                                    }
                                                    foreach (Parameter p in origFloor.Parameters)
                                                    {
                                                        try
                                                        {
                                                            Parameter newParam = flr.LookupParameter(p.Definition.Name);
                                                            if (newParam != null)
                                                            {
                                                                switch (newParam.StorageType)
                                                                {
                                                                    case StorageType.Double:
                                                                        newParam.Set(p.AsDouble());
                                                                        break;
                                                                    case StorageType.ElementId:
                                                                        newParam.Set(p.AsElementId());
                                                                        break;
                                                                    case StorageType.Integer:
                                                                        newParam.Set(p.AsInteger());
                                                                        break;
                                                                    case StorageType.String:
                                                                        newParam.Set(p.AsString());
                                                                        break;
                                                                    default:
                                                                        newParam.Set(p.AsString());
                                                                        break;
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Debug.WriteLine(ex.Message);
                                                        }
                                                    }

                                                    if (Math.Abs(offset - 0) > double.Epsilon)
                                                    {
                                                        Parameter p = flr.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                                                        p.Set(offset);
                                                    }
                                                    doc.Delete(origFloor.Id);

                                                    // Assign the parameters
                                                    SetParameters(flr, obj.Parameters, doc);

                                                    // Assign the GH InstanceGuid
                                                    AssignGuid(flr, uniqueId, instanceSchema, runId, nickName);
                                                }
                                                else // The curves coming in should match the floor sketch.  Let's modify the floor's locationcurves to edit it's location/shape
                                                {
                                                    try
                                                    {
                                                        int crvCount = 0;
                                                        foreach (ModelCurve l in mLines)
                                                        {
                                                            LocationCurve lc = l.Location as LocationCurve;
                                                            lc.Curve = crvArray.get_Item(crvCount);
                                                            crvCount++;
                                                        }

                                                        // Change the family and symbol if necessary
                                                        if (origFloor.FloorType.Name != floorType.Name)
                                                        {
                                                            try
                                                            {
                                                                origFloor.FloorType = floorType;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Debug.WriteLine(ex.Message);
                                                            }
                                                        }

                                                        // Set the incoming parameters
                                                        SetParameters(origFloor, obj.Parameters, doc);
                                                    }
                                                    catch // There was an error in trying to recreate it.  Just delete the original and recreate the thing.
                                                    {
                                                        flr = doc.Create.NewFloor(crvArray, floorType, lvl, false);

                                                        // Assign the parameters in the new floor to match the original floor object.
                                                        foreach (Parameter p in origFloor.Parameters)
                                                        {
                                                            try
                                                            {
                                                                Parameter newParam = flr.LookupParameter(p.Definition.Name);
                                                                if (newParam != null)
                                                                {
                                                                    switch (newParam.StorageType)
                                                                    {
                                                                        case StorageType.Double:
                                                                            newParam.Set(p.AsDouble());
                                                                            break;
                                                                        case StorageType.ElementId:
                                                                            newParam.Set(p.AsElementId());
                                                                            break;
                                                                        case StorageType.Integer:
                                                                            newParam.Set(p.AsInteger());
                                                                            break;
                                                                        case StorageType.String:
                                                                            newParam.Set(p.AsString());
                                                                            break;
                                                                        default:
                                                                            newParam.Set(p.AsString());
                                                                            break;
                                                                    }
                                                                }
                                                            }
                                                            catch (Exception exception)
                                                            {
                                                                Debug.WriteLine(exception.Message);
                                                            }
                                                        }

                                                        if (Math.Abs(offset - 0) > double.Epsilon)
                                                        {
                                                            Parameter p = flr.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                                                            p.Set(offset);
                                                        }

                                                        doc.Delete(origFloor.Id);

                                                        // Set the incoming parameters
                                                        SetParameters(flr, obj.Parameters, doc);
                                                        // Assign the GH InstanceGuid
                                                        AssignGuid(flr, uniqueId, instanceSchema, runId, nickName);
                                                    }
                                                }
                                            }
                                            else if (supressedModify) // Just modify the floor and don't risk replacing it.
                                            {
                                                flr = doc.GetElement(existingElems[i]) as Floor;

                                                // Change the family and symbol if necessary
                                                if (flr.FloorType.Name != floorType.Name)
                                                {
                                                    try
                                                    {
                                                        flr.FloorType = floorType;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                // Assign the parameters
                                                SetParameters(flr, obj.Parameters, doc);
                                            }
                                        }
                                        #endregion
                                        else if (obj.CategoryId == -2000035)
                                        {
                                            // Create a RoofExtrusion
                                            // Find the level
                                            Level lvl = GetLevel(obj.Curves[0].ControlPoints, doc);

                                            double offset = 0;
                                            if (Math.Abs(UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation) > double.Epsilon)
                                            {
                                                offset = UnitUtils.ConvertToInternalUnits(obj.Curves[0].ControlPoints[0].Z, lengthDUT) - lvl.ProjectElevation;
                                            }

                                            // Generate the curvearray from the incoming curves
                                            CurveArray crvArray = GetCurveArray(obj.Curves);

                                            // Create the roof
                                            FootPrintRoof roof = null;
                                            ModelCurveArray roofProfile = new ModelCurveArray();

                                            if (replace)  // Try to modify or create a new roof.
                                            {
                                                FootPrintRoof origRoof = doc.GetElement(existingElems[i]) as FootPrintRoof;

                                                // Find the model curves for the original wall
                                                ICollection<ElementId> ids;
                                                using (SubTransaction st = new SubTransaction(doc))
                                                {
                                                    st.Start();
                                                    ids = doc.Delete(origRoof.Id);
                                                    st.RollBack();
                                                }

                                                // Get the sketch curves for the roof object.
                                                List<ModelCurve> mLines = new List<ModelCurve>();
                                                foreach (ElementId id in ids)
                                                {
                                                    Element e = doc.GetElement(id);
                                                    if (e is ModelCurve)
                                                    {
                                                        mLines.Add(e as ModelCurve);
                                                    }
                                                }

                                                if (mLines.Count != crvArray.Size) // Sketch curves qty doesn't match up with the incoming cuves.  Just recreate the roof.
                                                {
                                                    roof = doc.Create.NewFootPrintRoof(crvArray, lvl, roofType, out roofProfile);

                                                    // Match parameters from the original roof to it's new iteration.
                                                    foreach (Parameter p in origRoof.Parameters)
                                                    {
                                                        try
                                                        {
                                                            Parameter newParam = roof.LookupParameter(p.Definition.Name);
                                                            if (newParam != null)
                                                            {
                                                                switch (newParam.StorageType)
                                                                {
                                                                    case StorageType.Double:
                                                                        newParam.Set(p.AsDouble());
                                                                        break;
                                                                    case StorageType.ElementId:
                                                                        newParam.Set(p.AsElementId());
                                                                        break;
                                                                    case StorageType.Integer:
                                                                        newParam.Set(p.AsInteger());
                                                                        break;
                                                                    case StorageType.String:
                                                                        newParam.Set(p.AsString());
                                                                        break;
                                                                    default:
                                                                        newParam.Set(p.AsString());
                                                                        break;
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Debug.WriteLine(ex.Message);
                                                        }
                                                    }

                                                    if (Math.Abs(offset - 0) > double.Epsilon)
                                                    {
                                                        Parameter p = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                                                        p.Set(offset);
                                                    }

                                                    doc.Delete(origRoof.Id);

                                                    // Set the new parameters
                                                    SetParameters(roof, obj.Parameters, doc);

                                                    // Assign the GH InstanceGuid
                                                    AssignGuid(roof, uniqueId, instanceSchema, runId, nickName);
                                                }
                                                else // The curves qty lines up, lets try to modify the roof sketch so we don't have to replace it.
                                                {
                                                    if (obj.CurveIds != null && obj.CurveIds.Count > 0)
                                                    {
                                                        // Just recreate the roof
                                                        roof = doc.Create.NewFootPrintRoof(crvArray, lvl, roofType, out roofProfile);

                                                        // Match parameters from the original roof to it's new iteration.
                                                        foreach (Parameter p in origRoof.Parameters)
                                                        {
                                                            try
                                                            {
                                                                Parameter newParam = roof.LookupParameter(p.Definition.Name);
                                                                if (newParam != null)
                                                                {
                                                                    switch (newParam.StorageType)
                                                                    {
                                                                        case StorageType.Double:
                                                                            newParam.Set(p.AsDouble());
                                                                            break;
                                                                        case StorageType.ElementId:
                                                                            newParam.Set(p.AsElementId());
                                                                            break;
                                                                        case StorageType.Integer:
                                                                            newParam.Set(p.AsInteger());
                                                                            break;
                                                                        case StorageType.String:
                                                                            newParam.Set(p.AsString());
                                                                            break;
                                                                        default:
                                                                            newParam.Set(p.AsString());
                                                                            break;
                                                                    }
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Debug.WriteLine(ex.Message);
                                                            }
                                                        }

                                                        if (Math.Abs(offset - 0) > double.Epsilon)
                                                        {
                                                            Parameter p = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                                                            p.Set(offset);
                                                        }

                                                        // Set the parameters from the incoming data
                                                        SetParameters(roof, obj.Parameters, doc);

                                                        // Assign the GH InstanceGuid
                                                        AssignGuid(roof, uniqueId, instanceSchema, runId, nickName);

                                                        doc.Delete(origRoof.Id);
                                                    }
                                                    else
                                                    {
                                                        try
                                                        {
                                                            int crvCount = 0;
                                                            foreach (ModelCurve l in mLines)
                                                            {
                                                                LocationCurve lc = l.Location as LocationCurve;
                                                                lc.Curve = crvArray.get_Item(crvCount);
                                                                crvCount++;
                                                            }

                                                            // Change the family and symbol if necessary
                                                            if (origRoof.RoofType.Name != roofType.Name)
                                                            {
                                                                try
                                                                {
                                                                    origRoof.RoofType = roofType;
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    Debug.WriteLine(ex.Message);
                                                                }
                                                            }

                                                            SetParameters(origRoof, obj.Parameters, doc);
                                                        }
                                                        catch // Modificaiton failed, lets just create a new roof.
                                                        {
                                                            roof = doc.Create.NewFootPrintRoof(crvArray, lvl, roofType, out roofProfile);

                                                            // Match parameters from the original roof to it's new iteration.
                                                            foreach (Parameter p in origRoof.Parameters)
                                                            {
                                                                try
                                                                {
                                                                    Parameter newParam = roof.LookupParameter(p.Definition.Name);
                                                                    if (newParam != null)
                                                                    {
                                                                        switch (newParam.StorageType)
                                                                        {
                                                                            case StorageType.Double:
                                                                                newParam.Set(p.AsDouble());
                                                                                break;
                                                                            case StorageType.ElementId:
                                                                                newParam.Set(p.AsElementId());
                                                                                break;
                                                                            case StorageType.Integer:
                                                                                newParam.Set(p.AsInteger());
                                                                                break;
                                                                            case StorageType.String:
                                                                                newParam.Set(p.AsString());
                                                                                break;
                                                                            default:
                                                                                newParam.Set(p.AsString());
                                                                                break;
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    Debug.WriteLine(ex.Message);
                                                                }
                                                            }

                                                            if (Math.Abs(offset - 0) > double.Epsilon)
                                                            {
                                                                Parameter p = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                                                                p.Set(offset);
                                                            }

                                                            // Set the parameters from the incoming data
                                                            SetParameters(roof, obj.Parameters, doc);

                                                            // Assign the GH InstanceGuid
                                                            AssignGuid(roof, uniqueId, instanceSchema, runId, nickName);

                                                            doc.Delete(origRoof.Id);
                                                        }
                                                    }
                                                }
                                            }
                                            else if (supressedModify) // Only update the parameters
                                            {
                                                roof = doc.GetElement(existingElems[i]) as FootPrintRoof;

                                                // Change the family and symbol if necessary
                                                if (roof.RoofType.Name != roofType.Name)
                                                {
                                                    try
                                                    {
                                                        roof.RoofType = roofType;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Debug.WriteLine(ex.Message);
                                                    }
                                                }
                                                // Assign the parameters
                                                SetParameters(roof, obj.Parameters, doc);
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Error", ex.Message);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                        t.Commit();
                    }
                }


            }
            #endregion

            //return succeeded;
        }

        private List<ElementId> FindExisting(Document doc, Guid uniqueId, int categoryId, int runid)
        {
            // Find existing elements with a matching GUID from the GH component.
            List<ElementId> existingElems = new List<ElementId>();

            Schema instanceSchema = Schema.Lookup(instanceSchemaGUID);
            if (instanceSchema == null)
            {
                return existingElems;
            }

            // find the existing elements
            if (categoryId == -2000011)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Walls);
                collector.OfClass(typeof(Wall));
                foreach (Wall w in collector)
                {
                    try
                    {
                        Entity entity = w.GetEntity(instanceSchema);
                        if (entity.IsValid())
                        {
                            Field f = instanceSchema.GetField("InstanceID");
                            string tempId = entity.Get<string>(f);
                            if (tempId == uniqueId.ToString())
                            {
                                if (runid == -1)
                                {
                                    existingElems.Add(w.Id);
                                }
                                else
                                {
                                    f = instanceSchema.GetField("RunID");
                                    int id = entity.Get<int>(f);
                                    if (id == runid)
                                    {
                                        existingElems.Add(w.Id);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
            else if (categoryId == -2000032)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Floors);
                collector.OfClass(typeof(Floor));
                foreach (Floor flr in collector)
                {
                    try
                    {
                        Entity entity = flr.GetEntity(instanceSchema);
                        if (entity.IsValid())
                        {
                            Field f = instanceSchema.GetField("InstanceID");
                            string tempId = entity.Get<string>(f);
                            if (tempId == uniqueId.ToString())
                            {
                                if (runid == -1)
                                {
                                    existingElems.Add(flr.Id);
                                }
                                else
                                {
                                    f = instanceSchema.GetField("RunID");
                                    int id = entity.Get<int>(f);
                                    if (id == runid)
                                    {
                                        existingElems.Add(flr.Id);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
            else if (categoryId == -2000035)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Roofs);
                collector.OfClass(typeof(FootPrintRoof));
                foreach (FootPrintRoof r in collector)
                {
                    try
                    {
                        Entity entity = r.GetEntity(instanceSchema);
                        if (entity.IsValid())
                        {
                            Field f = instanceSchema.GetField("InstanceID");
                            string tempId = entity.Get<string>(f);
                            if (tempId == uniqueId.ToString())
                            {
                                if (runid == -1)
                                {
                                    existingElems.Add(r.Id);
                                }
                                else
                                {
                                    f = instanceSchema.GetField("RunID");
                                    int id = entity.Get<int>(f);
                                    if (id == runid)
                                    {
                                        existingElems.Add(r.Id);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
            else if (categoryId == -2000240)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Levels);
                collector.OfClass(typeof(Level));

                foreach (Level l in collector)
                {
                    try
                    {
                        Entity entity = l.GetEntity(instanceSchema);
                        if (entity.IsValid())
                        {
                            Field f = instanceSchema.GetField("InstanceID");
                            string tempId = entity.Get<string>(f);
                            if (tempId == uniqueId.ToString())
                            {
                                if (runId == -1)
                                {
                                    existingElems.Add(l.Id);
                                }
                                else
                                {
                                    f = instanceSchema.GetField("RunID");
                                    int id = entity.Get<int>(f);
                                    if (id == runId)
                                    {
                                        existingElems.Add(l.Id);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
            else if (categoryId == -2000220)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_Grids);
                collector.OfClass(typeof(Grid));

                foreach (Grid g in collector)
                {
                    try
                    {
                        Entity entity = g.GetEntity(instanceSchema);
                        if (entity.IsValid())
                        {
                            Field f = instanceSchema.GetField("InstanceID");
                            string tempId = entity.Get<string>(f);
                            if (tempId == uniqueId.ToString())
                            {
                                if (runId == -1)
                                {
                                    existingElems.Add(g.Id);
                                }
                                else
                                {
                                    f = instanceSchema.GetField("RunID");
                                    int id = entity.Get<int>(f);
                                    if (id == runId)
                                    {
                                        existingElems.Add(g.Id);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
            else
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(FamilyInstance));

                foreach (Element e in collector)
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
                                if (tempId == uniqueId.ToString())
                                {
                                    if (runid == -1)
                                    {
                                        existingElems.Add(fi.Id);
                                    }
                                    else
                                    {
                                        f = instanceSchema.GetField("RunID");
                                        int id = entity.Get<int>(f);
                                        if (id == runid)
                                        {
                                            existingElems.Add(fi.Id);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
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

        private FamilySymbol FindFamilySymbol(string familyName, string typeName, Document doc)
        {
            FilteredElementCollector famCollector = new FilteredElementCollector(doc);
            famCollector.OfClass(typeof(Family));

            foreach (Family f in famCollector)
            {
                if (f.Name == familyName)
                {
                    foreach (ElementId fsid in f.GetFamilySymbolIds())
                    {
                        FamilySymbol fs = doc.GetElement(fsid) as FamilySymbol;
                        if (fs.Name == typeName)
                        {
                            FamilySymbol symbol = fs;
                            return symbol;
                        }
                    }
                }
            }
            return null;
        }

        private LevelType FindLevelType(string typeName, Document doc)
        {
            FilteredElementCollector ltCollector = new FilteredElementCollector(doc);
            ltCollector.OfClass(typeof(LevelType));

            foreach (LevelType lt in ltCollector)
            {
                if (lt.Name == typeName)
                    return lt;
            }
            return null;
        }

        private GridType FindGridType(string typeName, Document doc)
        {
            FilteredElementCollector gtCollector = new FilteredElementCollector(doc);
            gtCollector.OfClass(typeof(GridType));

            foreach (GridType gt in gtCollector)
            {
                if (gt.Name == typeName)
                    return gt;
            }
            return null;
        }

        private CurveArray GetCurveArray(IEnumerable<LyrebirdCurve> curves)
        {
            CurveArray crvArray = new CurveArray();
            int i = 0;
            foreach (LyrebirdCurve lbc in curves)
            {
                if (lbc.CurveType == "Circle")
                {
                    XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                    XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                    XYZ pt4 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[3].Z, lengthDUT));
                    XYZ pt5 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[4].Z, lengthDUT));
                    Arc arc1 = Arc.Create(pt1, pt3, pt2);
                    Arc arc2 = Arc.Create(pt3, pt5, pt4);
                    crvArray.Append(arc1);
                    crvArray.Append(arc2);
                }
                else if (lbc.CurveType == "Arc")
                {
                    XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                    XYZ pt3 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[2].Z, lengthDUT));
                    Arc arc = Arc.Create(pt1, pt3, pt2);
                    crvArray.Append(arc);
                }
                else if (lbc.CurveType == "Line")
                {
                    XYZ pt1 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[0].Z, lengthDUT));
                    XYZ pt2 = new XYZ(UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].X, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lbc.ControlPoints[1].Z, lengthDUT));
                    Line line = Line.CreateBound(pt1, pt2);
                    crvArray.Append(line);
                }
                else if (lbc.CurveType == "Spline")
                {
                    List<XYZ> controlPoints = new List<XYZ>();
                    List<double> weights = lbc.Weights;
                    List<double> knots = lbc.Knots;

                    foreach (LyrebirdPoint lp in lbc.ControlPoints)
                    {
                        XYZ pt = new XYZ(UnitUtils.ConvertToInternalUnits(lp.X, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Y, lengthDUT), UnitUtils.ConvertToInternalUnits(lp.Z, lengthDUT));
                        controlPoints.Add(pt);
                    }
                    try
                    {
                        if (lbc.Degree == 3)
                        {
                            NurbSpline spline = NurbSpline.Create(controlPoints, weights, knots, lbc.Degree, false, true);
                            crvArray.Append(spline);
                        }
                        else
                        {
                            HermiteSpline spline = HermiteSpline.Create(controlPoints, false);
                            crvArray.Append(spline);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                }
                i++;
            }
            return crvArray;
        }

        private Level GetLevel(List<LyrebirdPoint> controlPoints, Document doc)
        {
            Level lvl = null;

            FilteredElementCollector lvlCollector = new FilteredElementCollector(doc);
            lvlCollector.OfCategory(BuiltInCategory.OST_Levels);
            lvlCollector.OfClass(typeof(Level));
            foreach (Level l in lvlCollector)
            {
                try
                {
                    if (lvl == null)
                    {
                        lvl = l;
                    }
                    else
                    {
                        if (Math.Abs(l.Elevation - UnitUtils.ConvertToInternalUnits(controlPoints[0].Z, lengthDUT)) < Math.Abs(lvl.ProjectElevation - UnitUtils.ConvertToInternalUnits(controlPoints[0].Z, lengthDUT)))
                        {
                            lvl = l;
                        }
                    }
                }
                catch { }
            }

            return lvl;
        }

        private ElementId FindHost(XYZ location, int hostType, Document doc)
        {
            ElementId host = null;
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            // Subtransaction to insert a family and use it to check for intersctions.
            // The family is then moved around to check for the host of each new object being created
            // After the element creation process is over the object and it's parent family are deleted form the project.
            using (SubTransaction subTrans = new SubTransaction(doc))
            {
                subTrans.Start();

                if (hostFinder == null)
                {
                    // check if the point family exists
                    string path = new System.IO.FileInfo(typeof(LyrebirdService).Assembly.Location).DirectoryName + "\\IntersectionPoint.rfa";
                    //string path = typeof(LyrebirdService).Assembly.Location.Replace("LMNA.Lyrebird.Revit2015.dll", "IntersectionPoint.rfa");
                    if (!System.IO.File.Exists(path))
                    {
                        // save the file from this assembly and load it into project
                        //string directory = System.IO.Path.GetDirectoryName(path);
                        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        WriteResource(assembly, "IntersectionPoint.rfa", path);
                    }

                    // Load the family and place an instance of it.
                    Family insertPoint = null;
                    
                    try
                    {
                        if (System.IO.File.Exists(path))
                        {
                            doc.LoadFamily(path, out insertPoint);
                        }
                        else
                        {
                            TaskDialog.Show("Error", "Could not find family to load");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Error", ex.Message);
                    }
                    
                    if (insertPoint != null)
                    {
                        FamilySymbol ips = null;
                        foreach (ElementId fsid in insertPoint.GetFamilySymbolIds())
                        {
                            ips = doc.GetElement(fsid) as FamilySymbol;
                        }

                        // Create an instance
                        hostFinder = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(doc, ips);
                        System.IO.File.Delete(path);
                    }
                    else
                    {
                        TaskDialog.Show("test", "InsertPoint family is still null, loading didn't work.");
                    }
                }
                
                IList<ElementId> placePointIds = new List<ElementId>();
                placePointIds = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(hostFinder);
                try
                {
                    ReferencePoint rp = doc.GetElement(placePointIds[0]) as ReferencePoint;
                    XYZ movedPt;
                    if (hostType == 1 || hostType == 3)
                    {
                        movedPt = new XYZ(location.X, location.Y, location.Z + 0.00328);
                    }
                    else
                    {
                        movedPt = new XYZ(location.X, location.Y, location.Z - 0.00328);
                    }

                    if (rp != null)
                    {
                        XYZ vector = movedPt.Subtract(rp.Position);
                        ElementTransformUtils.MoveElement(doc, rp.Id, vector);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }

                Element elem = hostFinder as Element;
                if (elem != null)
                {
                    // Find the host element
                    if (hostType == 1)
                    {
                        // find a wall
                        collector.OfCategory(BuiltInCategory.OST_Walls);
                        collector.OfClass(typeof(Wall));

                        ElementIntersectsElementFilter intersectionFilter = new ElementIntersectsElementFilter(elem);
                        collector.WherePasses(intersectionFilter);
                        foreach (Element e in collector)
                        {
                            host = e.Id;
                        }
                    }
                    else if (hostType == 2)
                    {
                        // Find a floor
                        collector.OfCategory(BuiltInCategory.OST_Floors);
                        collector.OfClass(typeof(Floor));

                        ElementIntersectsElementFilter intersectionFilter = new ElementIntersectsElementFilter(elem);
                        collector.WherePasses(intersectionFilter);

                        foreach (Element e in collector)
                        {
                            host = e.Id;
                        }
                    }
                    else if (hostType == 3)
                    {
                        // find a ceiling
                        collector.OfCategory(BuiltInCategory.OST_Ceilings);
                        collector.OfClass(typeof(Ceiling));

                        ElementIntersectsElementFilter intersectionFilter = new ElementIntersectsElementFilter(elem);
                        collector.WherePasses(intersectionFilter);

                        foreach (Element e in collector)
                        {
                            host = e.Id;
                        }
                    }
                    else if (hostType == 4)
                    {
                        // find a roof
                        collector.OfCategory(BuiltInCategory.OST_Roofs);
                        collector.OfClass(typeof(RoofBase));

                        ElementIntersectsElementFilter intersectionFilter = new ElementIntersectsElementFilter(elem);
                        collector.WherePasses(intersectionFilter);

                        foreach (Element e in collector)
                        {
                            host = e.Id;
                        }
                    }
                }
                subTrans.Commit();

                // Delete the family file

            }

            return host;
        }


        private Face FindFace(XYZ location, XYZ vector, Document doc)
        {
            Face face = null;

            FilteredElementCollector elementCollector = new FilteredElementCollector(doc);
            elementCollector.WhereElementIsNotElementType();

            Options opt = new Options { ComputeReferences = true };
            foreach (Element e in elementCollector)
            {
                try
                {
                    GeometryElement ge = e.get_Geometry(opt);
                    foreach (GeometryObject go in ge)
                    {
                        try
                        {
                            XYZ endPt = location + vector;
                            Line l = Line.CreateBound(location - vector, endPt);
                            Curve c = l as Curve;
                            double dist = 1000000;
                            GeometryInstance geoInst = go as GeometryInstance;
                            if (geoInst != null)
                            {
                                // Family instance
                                GeometryElement instGeometry = geoInst.GetInstanceGeometry();
                                foreach (GeometryObject o in instGeometry)
                                {
                                    Solid s = o as Solid;
                                    if (s != null)
                                        foreach (Face f in s.Faces)
                                        {
                                            IntersectionResultArray results;
                                            SetComparisonResult result = f.Intersect(c, out results);
                                            if (results != null)
                                            {
                                                foreach (IntersectionResult res in results)
                                                {
                                                    XYZ intersect = res.XYZPoint;
                                                    double tempDist = location.DistanceTo(intersect);
                                                    if (tempDist < dist)
                                                    {
                                                        dist = tempDist;
                                                        face = f;
                                                        if (dist < 0.05)
                                                        {
                                                            return f;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                }
                            }
                            else
                            {
                                // Assume something like a wall.
                                Solid solid = go as Solid;

                                if (solid != null)
                                    foreach (Face f in solid.Faces)
                                    {
                                        IntersectionResultArray results;
                                        SetComparisonResult result = f.Intersect(c, out results);
                                        if (results != null)
                                        {
                                            foreach (IntersectionResult res in results)
                                            {
                                                XYZ intersect = res.XYZPoint;
                                                double tempDist = location.DistanceTo(intersect);
                                                if (tempDist < dist)
                                                {
                                                    dist = tempDist;
                                                    face = f;
                                                    if (dist < 0.05)
                                                    {
                                                        return f;
                                                    }
                                                }
                                            }
                                        }
                                    }
                            }
                        }
                        catch (Exception ex)
                        {
                            //errors++;
                            Debug.WriteLine(ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //errors++;
                    Debug.WriteLine(ex.Message);
                }
            }
            if (face == null)
            {
                TaskDialog.Show("error", "no face found");
            }
            return face;
        }

        public static void WriteResource(System.Reflection.Assembly targetAssembly, string resourceName, string filePath)
        {
            string[] resources = targetAssembly.GetManifestResourceNames();
            List<string> resoruceList = resources.ToList();
            foreach (string s in resoruceList)
            {
                if(s.Contains("IntersectionPoint.rfa"))
                {
                    using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(s))
                    {
                        using (System.IO.FileStream fileStream = new System.IO.FileStream(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), resourceName), System.IO.FileMode.Create))
                        {
                            for (int i = 0; i < stream.Length; i++)
                            {
                                fileStream.WriteByte((byte)stream.ReadByte());
                            }
                            fileStream.Close();
                        }
                    }
                }
            }
        }

        #region Assign Parameter Values
        private void SetParameters(FamilyInstance fi, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = fi.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                            try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = fi.Symbol.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                               try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        private void SetParameters(Wall wall, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = wall.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                           try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = wall.WallType.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                                try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        private void SetParameters(Floor floor, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = floor.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                            try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = floor.FloorType.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                                try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        private void SetParameters(FootPrintRoof roof, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = roof.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                            try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = roof.RoofType.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                               try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        private ElementId GetMaterial(string value, Document doc)
        {
            ElementId eid = null;

            try
            {
                eid = new ElementId(Convert.ToInt32(value));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            if (eid == null)
            {
                // Get the materials
                FilteredElementCollector matCollector = new FilteredElementCollector(doc);
                matCollector.OfClass(typeof(Material));
                foreach (Material m in matCollector)
                {
                    if (m.Name.ToUpper() == value.ToUpper())
                    {
                        eid = m.Id;
                    }
                }
            }

            return eid ?? (eid = Material.Create(doc, value));
        }

        private void SetParameters(Level lvl, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = lvl.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                            try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = lvl.LevelType.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                                try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        private void SetParameters(Grid grid, IEnumerable<RevitParameter> parameters, Document doc)
        {
            foreach (RevitParameter rp in parameters)
            {
                try
                {
                    Parameter p = grid.LookupParameter(rp.ParameterName);
                    switch (rp.StorageType)
                    {
                        case "Double":
                            if (p.Definition.ParameterType == ParameterType.Area)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                            else if (p.Definition.ParameterType == ParameterType.Volume)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                            else if (p.Definition.ParameterType == ParameterType.Length)
                                p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                            else
                                p.Set(Convert.ToDouble(rp.Value));
                            break;
                        case "Integer":
                            p.Set(Convert.ToInt32(rp.Value));
                            break;
                        case "String":
                            p.Set(rp.Value);
                            break;
                        case "ElementId":
                            try
                            {
                                int idInt = Convert.ToInt32(rp.Value);
                                ElementId elemId = new ElementId(idInt);
                                Element elem = doc.GetElement(elemId);
                                if (elem != null)
                                {
                                    //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                    p.Set(elemId);
                                }
                            }
                            catch
                            {
                                try
                                {
                                    p.Set(p.Definition.ParameterType == ParameterType.Material
                                        ? GetMaterial(rp.Value, doc)
                                        : new ElementId(Convert.ToInt32(rp.Value)));
                                }
                                catch (Exception ex)
                                {
                                    //TaskDialog.Show(p.Definition.Name, ex.Message);
                                }
                            }
                            break;
                        default:
                            p.Set(rp.Value);
                            break;
                    }
                }
                catch
                {
                    try
                    {
                        Parameter p = grid.GridType.LookupParameter(rp.ParameterName);
                        switch (rp.StorageType)
                        {
                            case "Double":
                                if (p.Definition.ParameterType == ParameterType.Area)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), areaDUT));
                                else if (p.Definition.ParameterType == ParameterType.Volume)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), volumeDUT));
                                else if (p.Definition.ParameterType == ParameterType.Length)
                                    p.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(rp.Value), lengthDUT));
                                else
                                    p.Set(Convert.ToDouble(rp.Value));
                                break;
                            case "Integer":
                                p.Set(Convert.ToInt32(rp.Value));
                                break;
                            case "String":
                                p.Set(rp.Value);
                                break;
                            case "ElementId":
                                try
                                {
                                    int idInt = Convert.ToInt32(rp.Value);
                                    ElementId elemId = new ElementId(idInt);
                                    Element elem = doc.GetElement(elemId);
                                    if (elem != null)
                                    {
                                        //TaskDialog.Show("Test:", "Param: " + p.Definition.Name + "\nID: " + elemId.IntegerValue.ToString());
                                        p.Set(elemId);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        p.Set(p.Definition.ParameterType == ParameterType.Material
                                            ? GetMaterial(rp.Value, doc)
                                            : new ElementId(Convert.ToInt32(rp.Value)));
                                    }
                                    catch (Exception ex)
                                    {
                                        //TaskDialog.Show(p.Definition.Name, ex.Message);
                                    }
                                }
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
        }

        #endregion

        #region Assign the GUID
        private void AssignGuid(FamilyInstance fi, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = fi.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }
            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                fi.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        private void AssignGuid(Wall wall, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = wall.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }

            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                wall.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }

        }

        private void AssignGuid(Floor floor, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = floor.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }
            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                floor.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        private void AssignGuid(FootPrintRoof roof, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = roof.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }
            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                roof.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        private void AssignGuid(Level lvl, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = lvl.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }
            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                lvl.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        private void AssignGuid(Grid grid, Guid guid, Schema instanceSchema, int run, string nickName)
        {
            Entity entity = null;
            try
            {
                entity = grid.GetEntity(instanceSchema);
            }
            catch (Exception ex)
            {
                Debug.Write("Error", ex.Message);
            }
            try
            {
                if (!entity.IsValid())
                {
                    entity = new Entity(instanceSchema);
                }
                Field field = instanceSchema.GetField("InstanceID");
                entity.Set<string>(field, guid.ToString());
                field = instanceSchema.GetField("RunID");
                entity.Set<int>(field, run);
                field = instanceSchema.GetField("NickName");
                entity.Set<string>(field, nickName);
                grid.SetEntity(entity);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }
        #endregion
    }
}
