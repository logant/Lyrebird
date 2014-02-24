using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using LMNA.Lyrebird.LyrebirdCommon;
using System.Diagnostics;

namespace LMNA.Lyrebird.GH
{
    #region ComponentInfo
    public class LyrebirdAssemblyPriority : GH_AssemblyPriority
    {
        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryShortName("LMNts", "LMN");
            Instances.ComponentServer.AddCategorySymbolName("LMNts", 'L');
            Instances.ComponentServer.AddCategoryIcon("LMNts", Properties.Resources.LMNts_16x16);

            return GH_LoadingInstruction.Proceed;
        }
    }

    public class LyrebirdInfo : GH_AssemblyInfo
    {
        public override string Version
        {
            get { return "1.1.0.0"; }
        }

        public override string AuthorName
        {
            get { return "LMN Tech Studio"; }
        }

        public override string AuthorContact
        {
            get { return "http://lmnts.lmnarchitects.com"; }
        }
    }
    #endregion

    public class DoubleClicker : Grasshopper.Kernel.Attributes.GH_ComponentAttributes
    {
        public DoubleClicker(GHClient comp)
            : base(comp)
        {
        }

        public override Grasshopper.GUI.Canvas.GH_ObjectResponse RespondToMouseDoubleClick(Grasshopper.GUI.Canvas.GH_Canvas sender, Grasshopper.GUI.GH_CanvasMouseEvent e)
        {
            ((GHClient)Owner).DisplayForm();
            return Grasshopper.GUI.Canvas.GH_ObjectResponse.Handled;
        }
    }

    public class GHClient : GH_Component
    {
        string objMessage = "Nothing has happened";
        string message = null;
        private List<RevitParameter> inputParameters = new List<RevitParameter>();
        private string familyName = "Not Selected";
        private string typeName = "Not Selected";
        private string category = "Not Selected";
        private int categoryId = -1;
        private List<LyrebirdId> uniqueIDs = new List<LyrebirdId>();
        int appVersion = 1;

        bool paramNamesEnabled = true;
        bool r2014 = true;
        bool r2015 = false;

        public List<RevitParameter> InputParams
        {
            get { return inputParameters; }
            set { inputParameters = value; }
        }

        public string FamilyName
        {
            get { return familyName; }
            set { familyName = value; }
        }

        public string TypeName
        {
            get { return typeName; }
            set { typeName = value; }
        }

        public new string Category
        {
            get { return category; }
            set { category = value; }
        }

        public int CategoryId
        {
            get { return categoryId; }
            set { categoryId = value; }
        }

        public List<LyrebirdId> UniqueIds
        {
            get { return uniqueIDs; }
            set { uniqueIDs = value; }
        }

        public GHClient()
            : base("Lyrebird Out", "LBOut", "Send data from GH to another application", "LMNts", "Utilities")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Trigger to stream the data from Grasshopper to another application.", GH_ParamAccess.item, false);
            pManager.AddPointParameter("Origin Point", "OP", "Origin points for sent objects.", GH_ParamAccess.tree, null);
            pManager.AddPointParameter("Adaptive Points", "AP", "Adaptive component points.", GH_ParamAccess.tree, null);
            pManager.AddCurveParameter("Curve", "C", "Single arc, line, or closed planar curves.  Closed planar curves can be used to generate floor, wall or roof sketches, or single segment non-closed arcs or lines can be used for line based family generation.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation", "O", "Vector to orient objects.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation on Face", "F", "Orientation of the element in relation to the face it will be hosted to", GH_ParamAccess.tree);
            pManager[1].Optional = true;
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
            pManager[5].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Selected Object", "Obj", "Object information that Lyrebird will create or modify", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "Msg", "Failure or warning messages", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool runCommand = false;
            GH_Structure<GH_Point> origPoints = new GH_Structure<GH_Point>();
            GH_Structure<GH_Point> adaptPoints = new GH_Structure<GH_Point>();
            GH_Structure<GH_Curve> curves = new GH_Structure<GH_Curve>();
            GH_Structure<GH_Vector> orientations = new GH_Structure<GH_Vector>();
            GH_Structure<GH_Vector> faceOrientations = new GH_Structure<GH_Vector>();
            DA.GetData(0, ref runCommand);
            DA.GetDataTree(1, out origPoints);
            DA.GetDataTree(2, out adaptPoints);
            DA.GetDataTree(3, out curves);
            DA.GetDataTree(4, out orientations);
            DA.GetDataTree(5, out faceOrientations);

            // Make sure the family and type is set before running the command.
            if (runCommand && (familyName == null || familyName == "Not Selected"))
            {
                message = "Please select a family/type by double-clicking on the component before running the command.";
            }
            else if (runCommand)
            {
                // Get the scale
                GHInfo ghi = new GHInfo();
                GHScale scale = ghi.GetScale(Rhino.RhinoDoc.ActiveDoc);

                // Send to Revit
                LyrebirdChannel channel = new LyrebirdChannel(appVersion);
                channel.Create();

                if (channel != null)
                {
                    string documentName = channel.DocumentName();
                    if (documentName != null)
                    {
                        // Create RevitObjects
                        List<RevitObject> obj = new List<RevitObject>();

                        #region OriginPoint Based
                        if (origPoints != null && origPoints.Branches.Count > 0)
                        {
                            List<RevitObject> tempObjs = new List<RevitObject>();
                            // make sure the branches match the datacount
                            if (origPoints.Branches.Count == origPoints.DataCount)
                            {
                                for (int i = 0; i < origPoints.Branches.Count; i++)
                                {
                                    GH_Point ghpt = origPoints[i][0];
                                    LyrebirdPoint point = new LyrebirdPoint
                                    {
                                        X = ghpt.Value.X,
                                        Y = ghpt.Value.Y,
                                        Z = ghpt.Value.Z
                                    };

                                    RevitObject ro = new RevitObject
                                    {
                                        Origin = point,
                                        FamilyName = familyName,
                                        TypeName = typeName,
                                        Category = category,
                                        CategoryId = categoryId,
                                        GHPath = origPoints.Paths[i].ToString(),
                                        GHScaleFactor = scale.ScaleFactor,
                                        GHScaleName = scale.ScaleName
                                    };

                                    tempObjs.Add(ro);
                                }
                                obj = tempObjs;
                            }
                            else
                            {
                                // Inform the user they need to graft their inputs.  Only one point per branch
                                System.Windows.Forms.MessageBox.Show("Warning:\n\nEach Branch represents an object, " +
                                    "so origin point based elements should be grafted so that each point is on it's own branch.");
                            }
                        }
                        #endregion

                        #region AdaptiveComponents
                        else if (adaptPoints != null && adaptPoints.Branches.Count > 0)
                        {
                            // generate adaptive components
                            List<RevitObject> tempObjs = new List<RevitObject>();
                            for (int i = 0; i < adaptPoints.Branches.Count; i++)
                            {
                                RevitObject ro = new RevitObject();
                                List<LyrebirdPoint> points = new List<LyrebirdPoint>();
                                for (int j = 0; j < adaptPoints.Branches[i].Count; j++)
                                {
                                    LyrebirdPoint point = new LyrebirdPoint(adaptPoints.Branches[i][j].Value.X, adaptPoints.Branches[i][j].Value.Y, adaptPoints.Branches[i][j].Value.Z);
                                    points.Add(point);
                                }
                                ro.AdaptivePoints = points;
                                ro.FamilyName = familyName;
                                ro.TypeName = typeName;
                                ro.Origin = null;
                                ro.Category = category;
                                ro.CategoryId = categoryId;
                                ro.GHPath = adaptPoints.Paths[i].ToString();
                                ro.GHScaleFactor = scale.ScaleFactor;
                                ro.GHScaleName = scale.ScaleName;
                                tempObjs.Add(ro);
                            }
                            obj = tempObjs;

                        }
                        #endregion

                        #region Curve Based
                        else if (curves != null && curves.Branches.Count > 0)
                        {
                            // Get curves for curve based components

                            // Determine if we're profile or line based
                            if (curves.Branches.Count == curves.DataCount)
                            {
                                // Determine if the curve is a closed planar curve
                                Curve tempCrv = curves.Branches[0][0].Value;
                                if (tempCrv.IsPlanar(0.00000001) && tempCrv.IsClosed)
                                {
                                    // Closed planar curve
                                    List<RevitObject> tempObjs = new List<RevitObject>();
                                    for (int i = 0; i < curves.Branches.Count; i++)
                                    {
                                        Curve crv = curves[i][0].Value;
                                        List<Curve> rCurves = new List<Curve>();
                                        bool getCrvs = CurveSegments(rCurves, crv, true);
                                        if (rCurves.Count > 0)
                                        {
                                            RevitObject ro = new RevitObject();
                                            List<LyrebirdCurve> lbCurves = new List<LyrebirdCurve>();
                                            for (int j = 0; j < rCurves.Count; j++)
                                            {
                                                LyrebirdCurve lbc;
                                                lbc = GetLBCurve(rCurves[j]);
                                                lbCurves.Add(lbc);
                                            }
                                            ro.Curves = lbCurves;
                                            ro.FamilyName = familyName;
                                            ro.Category = category;
                                            ro.CategoryId = categoryId;
                                            ro.TypeName = typeName;
                                            ro.Origin = null;
                                            ro.GHPath = curves.Paths[i].ToString();
                                            ro.GHScaleFactor = scale.ScaleFactor;
                                            ro.GHScaleName = scale.ScaleName;
                                            tempObjs.Add(ro);
                                        }
                                    }
                                    obj = tempObjs;

                                }
                                else if (!tempCrv.IsClosed)
                                {
                                    // Line based.  Can only be arc or linear curves
                                    List<RevitObject> tempObjs = new List<RevitObject>();
                                    for (int i = 0; i < curves.Branches.Count; i++)
                                    {

                                        Curve ghc = curves.Branches[i][0].Value;
                                        // Test that there is only one curve segment
                                        PolyCurve polycurve = ghc as PolyCurve;
                                        if (polycurve != null)
                                        {
                                            Curve[] segments = polycurve.Explode();
                                            if (segments.Count() != 1)
                                            {
                                                break;
                                            }
                                        }
                                        if (ghc != null)
                                        {
                                            //List<LyrebirdPoint> points = new List<LyrebirdPoint>();
                                            LyrebirdCurve lbc = GetLBCurve(ghc);
                                            List<LyrebirdCurve> lbcurves = new List<LyrebirdCurve> { lbc };

                                            RevitObject ro = new RevitObject
                                            {
                                                Curves = lbcurves,
                                                FamilyName = familyName,
                                                Category = category,
                                                CategoryId = categoryId,
                                                TypeName = typeName,
                                                Origin = null,
                                                GHPath = curves.Paths[i].ToString(),
                                                GHScaleFactor = scale.ScaleFactor,
                                                GHScaleName = scale.ScaleName
                                            };

                                            tempObjs.Add(ro);
                                        }
                                    }
                                    obj = tempObjs;
                                }
                            }
                            else
                            {
                                // Make sure all of the curves in each branch are closed
                                bool allClosed = true;
                                DataTree<CurveCheck> crvTree = new DataTree<CurveCheck>();
                                for (int i = 0; i < curves.Branches.Count; i++)
                                {
                                    List<GH_Curve> ghCrvs = curves.Branches[i];
                                    List<CurveCheck> checkedcurves = new List<CurveCheck>();
                                    GH_Path path = new GH_Path(i);
                                    for (int j = 0; j < ghCrvs.Count; j++)
                                    {
                                        Curve c = ghCrvs[j].Value;
                                        if (c.IsClosed)
                                        {
                                            AreaMassProperties amp = AreaMassProperties.Compute(c);
                                            if (amp != null)
                                            {
                                                double area = amp.Area;
                                                CurveCheck cc = new CurveCheck(c, area);
                                                checkedcurves.Add(cc);
                                            }
                                        }
                                        else
                                        {
                                            allClosed = false;
                                        }
                                    }
                                    if (allClosed)
                                    {
                                        // Sort the curves by area
                                        checkedcurves.Sort((x, y) => x.Area.CompareTo(y.Area));
                                        checkedcurves.Reverse();
                                        foreach (CurveCheck cc in checkedcurves)
                                        {
                                            crvTree.Add(cc, path);
                                        }
                                    }
                                }

                                if (allClosed)
                                {
                                    // Determine if the smaller profiles are within the larger
                                    bool allInterior = true;
                                    List<RevitObject> tempObjs = new List<RevitObject>();
                                    for (int i = 0; i < crvTree.Branches.Count; i++)
                                    {
                                        try
                                        {
                                            List<int> crvSegmentIds = new List<int>();
                                            List<LyrebirdCurve> lbCurves = new List<LyrebirdCurve>();
                                            List<CurveCheck> checkedCrvs = crvTree.Branches[i];
                                            Curve outerProfile = checkedCrvs[0].Curve;
                                            double outerArea = checkedCrvs[0].Area;
                                            List<Curve> planarCurves = new List<Curve>();
                                            planarCurves.Add(outerProfile);
                                            double innerArea = 0.0;
                                            for (int j = 1; j < checkedCrvs.Count; j++)
                                            {
                                                planarCurves.Add(checkedCrvs[j].Curve);
                                                innerArea += checkedCrvs[j].Area;
                                            }
                                            // Try to create a planar surface
                                            IEnumerable<Curve> surfCurves = planarCurves;
                                            Brep[] b = Brep.CreatePlanarBreps(surfCurves);
                                            if (b.Count() == 1)
                                            {
                                                // Test the areas
                                                double brepArea = b[0].GetArea();
                                                double calcArea = outerArea - innerArea;
                                                double diff = (brepArea - calcArea) / calcArea;

                                                if (diff < 0.1)
                                                {
                                                    // The profiles probably are all interior
                                                    foreach (CurveCheck cc in checkedCrvs)
                                                    {
                                                        Curve c = cc.Curve;
                                                        List<Curve> rCurves = new List<Curve>();
                                                        bool getCrvs = CurveSegments(rCurves, c, true);

                                                        if (rCurves.Count > 0)
                                                        {
                                                            int crvSeg = rCurves.Count;
                                                            crvSegmentIds.Add(crvSeg);
                                                            foreach (Curve rc in rCurves)
                                                            {
                                                                LyrebirdCurve lbc;
                                                                lbc = GetLBCurve(rc);
                                                                lbCurves.Add(lbc);
                                                            }
                                                        }
                                                    }
                                                    RevitObject ro = new RevitObject();
                                                    ro.Curves = lbCurves;
                                                    ro.FamilyName = familyName;
                                                    ro.Category = category;
                                                    ro.CategoryId = categoryId;
                                                    ro.TypeName = typeName;
                                                    ro.Origin = null;
                                                    ro.GHPath = crvTree.Paths[i].ToString();
                                                    ro.GHScaleFactor = scale.ScaleFactor;
                                                    ro.GHScaleName = scale.ScaleName;
                                                    ro.CurveIds = crvSegmentIds;
                                                    tempObjs.Add(ro);
                                                }
                                            }
                                            else
                                            {
                                                allInterior = false;
                                                message = "Warning:\n\nEach Branch represents an object, " +
                                                "so curve based elements should be grafted so that each curve is on it's own branch, or all curves on a branch should " +
                                                "be interior to the largest, outer boundary.";
                                            }
                                        }
                                        catch
                                        {
                                            allInterior = false;
                                            // Inform the user they need to graft their inputs.  Only one curve per branch
                                            message = "Warning:\n\nEach Branch represents an object, " +
                                                "so curve based elements should be grafted so that each curve is on it's own branch, or all curves on a branch should " +
                                                "be interior to the largest, outer boundary.";
                                        }
                                    }
                                    if (tempObjs.Count > 0)
                                    {
                                        obj = tempObjs;
                                    }
                                }
                            }
                        }
                        #endregion

                        // Orientation
                        if (orientations != null && orientations.Branches.Count > 0)
                        {
                            List<RevitObject> tempList = AssignOrientation(obj, orientations);
                            obj = tempList;
                        }

                        // face orientation
                        if (faceOrientations != null && faceOrientations.Branches.Count > 0)
                        {
                            List<RevitObject> tempList = AssignFaceOrientation(obj, faceOrientations);
                            obj = tempList;
                        }

                        // Parameters...
                        if (Params.Input.Count > 6)
                        {
                            List<RevitObject> currentObjs = obj;
                            List<RevitObject> tempObjs = new List<RevitObject>();
                            for (int r = 0; r < currentObjs.Count; r++)
                            {
                                RevitObject ro = currentObjs[r];
                                List<RevitParameter> revitParams = new List<RevitParameter>();
                                for (int i = 6; i < Params.Input.Count; i++)
                                {

                                    RevitParameter rp = new RevitParameter();
                                    IGH_Param param = Params.Input[i];
                                    string paramInfo = param.Description;
                                    string[] pi = paramInfo.Split(new[] { "\n", ":" }, StringSplitOptions.None);
                                    string paramName = null;
                                    try
                                    {
                                        paramName = pi[1].Substring(1);
                                        string paramStorageType = pi[5].Substring(1);
                                        rp.ParameterName = paramName;
                                        rp.StorageType = paramStorageType;
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine(ex.Message);
                                    }
                                    if (paramName != null)
                                    {
                                        GH_Structure<IGH_Goo> data = null;
                                        try
                                        {
                                            DA.GetDataTree(i, out data);
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                        }
                                        if (data != null)
                                        {
                                            string value = data[r][0].ToString();
                                            rp.Value = value;
                                            revitParams.Add(rp);
                                        }
                                    }

                                }
                                ro.Parameters = revitParams;
                                tempObjs.Add(ro);
                            }
                            obj = tempObjs;
                        }

                        // Send the data to Revit to create and/or modify family instances.
                        if (obj != null && obj.Count > 0)
                        {
                            try
                            {
                                string docName = channel.DocumentName();
                                if (docName == null || docName == string.Empty)
                                {
                                    message = "Could not contact the lyrebird server.  Make sure it's running and try again.";
                                }
                                else
                                {
                                    string nn = NickName;
                                    if (nn == null || nn.Length == 0)
                                    {
                                        nn = "LBOut";
                                    }
                                    channel.CreateOrModify(obj, InstanceGuid, NickName);
                                    message = obj.Count.ToString() + " objects sent to the lyrebird server.";
                                }
                            }
                            catch
                            {
                                message = "Could not contact the lyrebird server.  Make sure it's running and try again.";
                            }
                        }
                        channel.Dispose();
                        try
                        {
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                    }
                    else
                    {
                        message = "Error\n" + "The Lyrebird Service could not be found.  Ensure Revit is running, the Lyrebird server plugin is installed, and the server is active.";
                    }
                }
            }
            else
            {
                message = null;
            }

            // Check if the revit information is set
            if (familyName != null || (familyName != "Not Selected" && typeName != "Not Selected"))
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Family: " + familyName);
                sb.AppendLine("Type: " + typeName);
                sb.AppendLine("Category: " + category);
                for (int i = 0; i < inputParameters.Count; i++)
                {
                    RevitParameter rp = inputParameters[i];
                    string type = "Instance";
                    if (rp.IsType)
                    {
                        type = "Type";
                    }
                    sb.AppendLine(string.Format("Parameter{0}: {1}  /  {2}  /  {3}", (i + 1).ToString(CultureInfo.InvariantCulture), rp.ParameterName, rp.StorageType, type));
                }
                objMessage = sb.ToString();
            }
            else
            {
                objMessage = "No data type set.  Double-click to set data type";
            }

            DA.SetData(0, objMessage);
            DA.SetData(1, message);
        }

        public override Guid ComponentGuid
        {
            get { return new Guid("f53a7976-fafe-44ec-8df4-de8261646383"); }
        }

        protected override System.Drawing.Bitmap Icon
        {
            get { return Properties.Resources.LyreBird_24x24; }
        }

        public void DisplayForm()
        {

            LyrebirdChannel channel = new LyrebirdChannel(appVersion);
            channel.Create();
            
            if (channel != null)
            {
                try
                {
                    SetRevitDataForm form = new SetRevitDataForm(channel, this);
                    form.ShowDialog();
                    if (form.DialogResult.HasValue && form.DialogResult.Value)
                    {
                        ExpireSolution(true);
                        SyncInputs();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    //System.Windows.Forms.MessageBox.Show("The Lyrebird Service could not be found.  Ensure Revit is running, the Lyrebird server plugin is installed, and the server is active.");
                }

                channel.Dispose();
            }
        }

        public override void CreateAttributes()
        {
            m_attributes = new DoubleClicker(this);
        }

        protected override void AppendAdditionalComponentMenuItems(System.Windows.Forms.ToolStripDropDown iMenu)
        {
            Menu_AppendItem(iMenu, "Full Parameter Names", Menu_ParamNamesClicked, true, paramNamesEnabled);
            
            // Application options.  Only 2014 will work for now.  Consider removing 2015 and/or adding support for 2013
            System.Windows.Forms.ToolStripMenuItem appItem = Menu_AppendItem(iMenu, "Application");
            appItem.DropDownItems.Add(Menu_AppendItem(iMenu, "Revit 2014", Menu_R2014Clicked, true, r2014));
            appItem.DropDownItems.Add(Menu_AppendItem(iMenu, "Revit 2015", Menu_R2015Clicked, true, r2015));
        }

        private void Menu_R2014Clicked(object sender, EventArgs e)
        {
            r2014 = !r2014;
            if (r2014)
            {
                r2015 = false;
            }
            appVersion = 1;
        }

        private void Menu_R2015Clicked(object sender, EventArgs e)
        {
            r2015 = !r2015;
            if (r2015)
            {
                r2014 = false;
            }
            appVersion = 2;
        }

        private void Menu_ParamNamesClicked(object sender, EventArgs e)
        {
            paramNamesEnabled = !paramNamesEnabled;
        }

        // Simple test to see if the communication is working  Just gets the project document name
        private void Menu_TestClicked(object sender, EventArgs e)
        {
            // Currently removed from the menu
            try
            {
                LyrebirdChannel channel = new LyrebirdChannel(appVersion);
                channel.Create();
                string test = channel.DocumentName() ?? "Failed to Collect Document Name";
                string write = "Document: " + test;
                System.Windows.Forms.MessageBox.Show(write);
                
                channel.Dispose();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // Originally this was done from the right-click menu.
        // It has since been changed to run whenever parameters are added/removed
        private void SyncInputs()
        {
            // Find out how many parameters there are and if inputs should be added or removed.
            if (inputParameters.Count == Params.Input.Count - 6)
            {
                // Parameters quantities match up with inputs, do nothing
                RefreshParameters();
                return;
            }
                
            RecordUndoEvent("Sync Inputs");

            //  Check if we need to add inputs
            if (inputParameters.Count > Params.Input.Count - 6)
            {
                for (int i = Params.Input.Count + 1; i <= inputParameters.Count + 6; i++)
                {
                    Grasshopper.Kernel.Parameters.Param_GenericObject param = new Grasshopper.Kernel.Parameters.Param_GenericObject
                    {
                        Name = "Parameter" + (i - 6).ToString(CultureInfo.InvariantCulture),
                        NickName = "P" + (i - 6).ToString(CultureInfo.InvariantCulture),
                        Description =
                            "Parameter Name: " + inputParameters[i - 7].ParameterName + "\nIs Type: " +
                            inputParameters[i - 7].IsType.ToString() + "\nStorageType: " +
                            inputParameters[i - 7].StorageType,
                        Optional = true,
                        Access = GH_ParamAccess.tree
                    };
                        
                    Params.RegisterInputParam(param);
                }
            }

            // Remove unnecessay inputs
            else if (inputParameters.Count < Params.Input.Count - 6)
            {
                while (Params.Input.Count > inputParameters.Count + 6)
                {
                    IGH_Param param = Params.Input[Params.Input.Count - 1];
                    Params.UnregisterInputParameter(param);
                }
            }
                
            RefreshParameters();
            Params.OnParametersChanged();
            ExpireSolution(true);
        }

        private void RefreshParameters()
        {
            for (int i = 0; i < inputParameters.Count; i++)
            {
                try
                {
                    IGH_Param param = Params.Input[i + 6];
                    if (paramNamesEnabled)
                    {
                        param.NickName = inputParameters[i].ParameterName;
                    }
                    else
                    {
                        bool renamed = false;
                        try
                        {
                            int x = Convert.ToInt32(param.NickName.Substring(1, 1));
                        }
                        catch
                        {
                            renamed = true;
                        }
                        if (param.NickName.Length == 2 && param.NickName.Substring(0, 1) == "P" && !renamed)
                        {
                            param.NickName = "P" + (i + 1).ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    param.Name = "Parameter" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    //param.NickName = "P" + (i + 1).ToString();
                    param.Description = "Parameter Name: " + inputParameters[i].ParameterName + "\nIs Type: " + inputParameters[i].IsType.ToString() + "\nStorageType: " + inputParameters[i].StorageType;
                    Params.RegisterInputParam(param);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }

        public override bool Write(GH_IO.Serialization.GH_IWriter writer)
        {
            // Add the family name and type name
            writer.SetString("FamilyName", FamilyName);
            writer.SetString("TypeName", TypeName);
            writer.SetString("Category", Category);
            writer.SetInt32("CategoryId", CategoryId);
            for (int i = 0; i < inputParameters.Count; i++)
            {
                try
                {
                    RevitParameter rp = inputParameters[i];
                    writer.SetString("ParameterName" + i.ToString(CultureInfo.InvariantCulture), rp.ParameterName);
                    writer.SetString("StorageType" + i.ToString(CultureInfo.InvariantCulture), rp.StorageType);
                    writer.SetBoolean("IsType" + i.ToString(CultureInfo.InvariantCulture), rp.IsType);
                }
                catch (Exception exception)
                {
                  Debug.WriteLine(exception.Message);
                }
            }
            return base.Write(writer);
        }

        public override bool Read(GH_IO.Serialization.GH_IReader reader)
        {
            FamilyName = reader.GetString("FamilyName");
            TypeName = reader.GetString("TypeName");
            Category = reader.GetString("Category");
            CategoryId = reader.GetInt32("CategoryId");
            bool test = true;
            int i = 0;
            List<RevitParameter> parameters = new List<RevitParameter>();
            while (test)
            {
                RevitParameter rp = new RevitParameter();
                try
                {
                    rp.ParameterName = reader.GetString("ParameterName" + i.ToString(CultureInfo.InvariantCulture));
                    rp.StorageType = reader.GetString("StorageType" + i.ToString(CultureInfo.InvariantCulture));
                    rp.IsType = reader.GetBoolean("IsType" + i.ToString(CultureInfo.InvariantCulture));
                    parameters.Add(rp);
                }
                catch
                {
                    test = false;
                }
                i++;
            }

            InputParams = parameters;
            if (inputParameters.Count > 0)
            {
                SyncInputs();
            }
            return base.Read(reader);
        }

        private List<RevitObject> AssignOrientation(List<RevitObject> currentList, GH_Structure<GH_Vector> orientations)
        {
            List<RevitObject> current = currentList;
            List<RevitObject> tempObj = new List<RevitObject>();
            for (int i = 0; i < orientations.Branches.Count; i++)
            {
                RevitObject ro = current[i];
                try
                {
                    Vector3d v = orientations[i][0].Value;

                    LyrebirdPoint p = new LyrebirdPoint {X = v.X, Y = v.Y, Z = v.Z};
                    ro.Orientation = p;
                    tempObj.Add(ro);
                }
                catch (Exception exception)
                {
                  Debug.WriteLine(exception.Message);
                }
            }
            return tempObj;
        }

        private List<RevitObject> AssignFaceOrientation(List<RevitObject> currentList, GH_Structure<GH_Vector> orientations)
        {
            List<RevitObject> current = currentList;
            List<RevitObject> tempObj = new List<RevitObject>();
            for (int i = 0; i < orientations.Branches.Count; i++)
            {
                RevitObject ro = current[i];
                try
                {
                    Vector3d v = orientations[i][0].Value;

                    LyrebirdPoint p = new LyrebirdPoint {X = v.X, Y = v.Y, Z = v.Z};
                    ro.FaceOrientation = p;
                    tempObj.Add(ro);
                }
                catch (Exception exception)
                {
                  Debug.WriteLine(exception.Message);
                }
            }
            return tempObj;
        }

        private LyrebirdCurve GetLBCurve(Curve crv)
        {
            LyrebirdCurve lbc = null;
            
            List<LyrebirdPoint> points = new List<LyrebirdPoint>();
            if (crv.IsLinear())
            {
                // standard linear element
                points.Add(new LyrebirdPoint(crv.PointAtStart.X, crv.PointAtStart.Y, crv.PointAtStart.Z));
                points.Add(new LyrebirdPoint(crv.PointAtEnd.X, crv.PointAtEnd.Y, crv.PointAtEnd.Z));
                lbc = new LyrebirdCurve(points, "Line");
            }
            else if (crv.IsCircle())
            {
                crv.Domain = new Interval(0, 1);
                points.Add(new LyrebirdPoint(crv.PointAtStart.X, crv.PointAtStart.Y, crv.PointAtStart.Z));
                points.Add(new LyrebirdPoint(crv.PointAt(0.25).X, crv.PointAt(0.25).Y, crv.PointAt(0.25).Z));
                points.Add(new LyrebirdPoint(crv.PointAt(0.5).X, crv.PointAt(0.5).Y, crv.PointAt(0.5).Z));
                points.Add(new LyrebirdPoint(crv.PointAt(0.75).X, crv.PointAt(0.75).Y, crv.PointAt(0.75).Z));
                points.Add(new LyrebirdPoint(crv.PointAtEnd.X, crv.PointAtEnd.Y, crv.PointAtEnd.Z));
                lbc = new LyrebirdCurve(points, "Circle");
            }
            else if (crv.IsArc())
            {
                crv.Domain = new Interval(0, 1);
                // standard arc element
                points.Add(new LyrebirdPoint(crv.PointAtStart.X, crv.PointAtStart.Y, crv.PointAtStart.Z));
                points.Add(new LyrebirdPoint(crv.PointAt(0.5).X, crv.PointAt(0.5).Y, crv.PointAt(0.5).Z));
                points.Add(new LyrebirdPoint(crv.PointAtEnd.X, crv.PointAtEnd.Y, crv.PointAtEnd.Z));
                lbc = new LyrebirdCurve(points, "Arc");
            }
            else
            {
                // Spline
                // Old line: if (crv.Degree >= 3)
                if (crv.Degree == 3)
                {
                    NurbsCurve nc = crv as NurbsCurve;
                    if (nc != null)
                    {
                        List<LyrebirdPoint> lbPoints = new List<LyrebirdPoint>();
                        List<double> weights = new List<double>();
                        List<double> knots = new List<double>();

                        foreach (ControlPoint cp in nc.Points)
                        {
                            LyrebirdPoint pt = new LyrebirdPoint(cp.Location.X, cp.Location.Y, cp.Location.Z);
                            double weight = cp.Weight;
                            lbPoints.Add(pt);
                            weights.Add(weight);
                        }
                        for (int k = 0; k < nc.Knots.Count; k++)
                        {
                            double knot = nc.Knots[k];
                            // Add a duplicate knot for the first and last knot in the Rhino curve.
                            // Revit needs 2 more knots than Rhino to define a spline.
                            if (k == 0 || k == nc.Knots.Count - 1)
                            {
                                knots.Add(knot);
                            }
                            knots.Add(knot);
                        }

                        lbc = new LyrebirdCurve(lbPoints, weights, knots, nc.Degree, nc.IsPeriodic) { CurveType = "Spline" };
                    }
                }
                else
                {
                    const double incr = 1.0 / 100;
                    List<LyrebirdPoint> pts = new List<LyrebirdPoint>();
                    List<double> weights = new List<double>();
                    for (int i = 0; i <= 100; i++)
                    {
                        Point3d pt = crv.PointAtNormalizedLength(i * incr);
                        LyrebirdPoint lbp = new LyrebirdPoint(pt.X, pt.Y, pt.Z);
                        weights.Add(1.0);
                        pts.Add(lbp);
                    }

                    lbc = new LyrebirdCurve(pts, "Spline") { Weights = weights, Degree = crv.Degree };
                }
            }

            return lbc;
        }

        protected bool CurveSegments(List<Curve> L, Curve crv, bool recursive)
        {
            if (crv == null) { return false; }
            PolyCurve polycurve = crv as PolyCurve;
            if (polycurve != null)
            {
                if (recursive) { polycurve.RemoveNesting(); }
                Curve[] segments = polycurve.Explode();
                if (segments == null) { return false; }
                if (segments.Length == 0) { return false; }
                if (recursive)
                {
                    foreach (Curve S in segments)
                    {
                        CurveSegments(L, S, true);
                    }
                }
                else
                {
                    foreach (Curve S in segments)
                    {
                        L.Add(S.DuplicateShallow() as Curve);
                    }
                }
                return true;
            }
            PolylineCurve polyline = crv as PolylineCurve;
            if (polyline != null)
            {
                if (recursive)
                {
                    for (int i = 0; i < (polyline.PointCount - 1); i++)
                    {
                        L.Add(new LineCurve(polyline.Point(i), polyline.Point(i + 1)));
                    }
                }
                else
                {
                    L.Add(polyline.DuplicateCurve());
                }
                return true;
            }
            Polyline p;
            if (crv.TryGetPolyline(out p))
            {
                if (recursive)
                {
                    for (int i = 0; i < (p.Count - 1); i++)
                    {
                        L.Add(new LineCurve(p[i], p[i + 1]));
                    }
                }
                else
                {
                    L.Add(new PolylineCurve(p));
                }
                return true;
            }
            //Maybe it's a LineCurve?
            LineCurve line = crv as LineCurve;
            if (line != null)
            {
                L.Add(line.DuplicateCurve());
                return true;
            }
            //It might still be an ArcCurve...
            ArcCurve arc = crv as ArcCurve;
            if (arc != null)
            {
                L.Add(arc.DuplicateCurve());
                return true;
            }
            //Nothing else worked, lets assume it's a nurbs curve and go from there...
            NurbsCurve nurbs = crv.ToNurbsCurve();
            if (nurbs == null) { return false; }
            double t0 = nurbs.Domain.Min;
            double t1 = nurbs.Domain.Max;
            int LN = L.Count;
            do
            {
              double t;
              if (!nurbs.GetNextDiscontinuity(Continuity.C1_locus_continuous, t0, t1, out t)) { break; }
                Interval trim = new Interval(t0, t);
                if (trim.Length < 1e-10)
                {
                    t0 = t;
                    continue;
                }
                Curve M = nurbs.DuplicateCurve();
                M = M.Trim(trim);
                if (M.IsValid) { L.Add(M); }
                t0 = t;
            } while (true);
            if (L.Count == LN) { L.Add(nurbs); }
            return true;
        }
    }

    public class CurveCheck
    {
        public Curve Curve { get; set; }
        public double Area { get; set; }

        public CurveCheck(Curve crv, double area)
        {
            Curve = crv;
            Area = area;
        }
    }
}
