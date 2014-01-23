using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GH_IO;
using Grasshopper;
using Grasshopper.GUI;
using Grasshopper.Kernel;
using System.ServiceModel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

using LMNA.Lyrebird.LyrebirdCommon;

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
        public override string Description
        {
            get { return base.Description; }
        }

        public override string Name
        {
            get { return base.Name; }
        }

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
            ((GHClient)this.Owner).DisplayForm();
            return Grasshopper.GUI.Canvas.GH_ObjectResponse.Handled;
        }
    }

    public class GHClient : GH_Component
    {
        string message = "Nothing has happened";
        private List<RevitParameter> inputParameters = new List<RevitParameter>();
        private string familyName = "Not Selected";
        private string typeName = "Not Selected";
        private string category = "Not Selected";
        private List<LyrebirdId> uniqueIDs = new List<LyrebirdId>();
        int appVersion = 1;

        bool paramNamesEnabled = true;

        private string[] m_settings;

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

        public string Category
        {
            get { return category; }
            set { category = value; }
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

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Trigger to stream the data from Grasshopper to another application.", GH_ParamAccess.item, false);
            pManager.AddIntegerParameter("Application", "A", "Application that will receive the data.", GH_ParamAccess.item, 1);
            pManager.AddPointParameter("Origin Point", "OP", "Origin points for sent objects.", GH_ParamAccess.tree, null);
            pManager.AddPointParameter("Adaptive Points", "AP", "Adaptive component points.", GH_ParamAccess.tree, null);
            pManager.AddCurveParameter("Curve", "C", "Single arc, line, or closed planar curves.  Closed planar curves can be used to generate floor, wall or roof sketches, or single segment non-closed arcs or lines can be used for line based family generation.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation", "O", "Vector to orient objects.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation on Face", "F", "Orientation of the element in relation to the face it will be hosted to", GH_ParamAccess.tree);
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
            pManager[5].Optional = true;
            pManager[6].Optional = true;
            Grasshopper.Kernel.Parameters.Param_Integer paramInt = pManager[1] as Grasshopper.Kernel.Parameters.Param_Integer;
            if (paramInt != null)
            {
                paramInt.AddNamedValue("Send to Revit 2013", 0);
                paramInt.AddNamedValue("Send to Revit 2014", 1);
                paramInt.AddNamedValue("Send to Revit 2015", 2);
            }
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Msg", "Msg", "Temporary message", GH_ParamAccess.item);
            pManager.AddTextParameter("Guid", "G", "Guids for this component instance GH", GH_ParamAccess.item);

            // TODO: See about tracking ElementId's of created revit elements back to the GH component
            // Returning ID's from Revit doesn't work as is since it runs on a separate thread and this continues before it's finished
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
            DA.GetData(1, ref appVersion);
            DA.GetDataTree(2, out origPoints);
            DA.GetDataTree(3, out adaptPoints);
            DA.GetDataTree(4, out curves);
            DA.GetDataTree(5, out orientations);
            DA.GetDataTree(6, out faceOrientations);
            if (runCommand == true)
            {
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

                            for (int i = 0; i < origPoints.Branches.Count; i++)
                            {
                                GH_Point ghpt = origPoints[i][0];
                                LyrebirdPoint point = new LyrebirdPoint();
                                point.X = ghpt.Value.X;
                                point.Y = ghpt.Value.Y;
                                point.Z = ghpt.Value.Z;

                                RevitObject ro = new RevitObject();
                                ro.Origin = point;
                                ro.FamilyName = familyName;
                                ro.TypeName = typeName;
                                ro.Category = category;
                                ro.GHPath = origPoints.Paths[i].ToString();
                                tempObjs.Add(ro);
                            }
                            obj = tempObjs;
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
                                ro.GHPath = adaptPoints.Paths[i].ToString();
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
                                Rhino.Geometry.Curve tempCrv = curves.Branches[0][0].Value;
                                if (tempCrv.IsPlanar() && tempCrv.IsClosed)
                                {
                                    // Closed planar curve
                                    List<RevitObject> tempObjs = new List<RevitObject>();
                                    for (int i = 0; i < curves.Branches.Count; i++)
                                    {
                                        Rhino.Geometry.Curve crv = curves[i][0].Value;
                                        List<Rhino.Geometry.Curve> rCurves = new List<Rhino.Geometry.Curve>();
                                        bool getCrvs = CurveSegments(rCurves, crv, true);
                                        if (rCurves.Count > 0)
                                        {
                                            RevitObject ro = new RevitObject();
                                            List<LyrebirdCurve> lbCurves = new List<LyrebirdCurve>();
                                            for (int j = 0; j < rCurves.Count; j++)
                                            {
                                                LyrebirdCurve lbc = null;
                                                lbc = GetLBCurve(rCurves[j]);
                                                lbCurves.Add(lbc);
                                            }
                                            ro.Curves = lbCurves;
                                            ro.FamilyName = familyName;
                                            ro.Category = category;
                                            ro.TypeName = typeName;
                                            ro.Origin = null;
                                            ro.GHPath = curves.Paths[i].ToString();
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

                                        Rhino.Geometry.Curve ghc = curves.Branches[i][0].Value;
                                        // Test that there is only one curve segment
                                        Rhino.Geometry.PolyCurve polycurve = ghc as Rhino.Geometry.PolyCurve;
                                        if (polycurve != null)
                                        {
                                            Rhino.Geometry.Curve[] segments = polycurve.Explode();
                                            if (segments.Count() != 1)
                                            {
                                                break;
                                            }
                                        }
                                        if (ghc != null)
                                        {
                                            List<LyrebirdPoint> points = new List<LyrebirdPoint>();
                                            LyrebirdCurve lbc = GetLBCurve(ghc);
                                            List<LyrebirdCurve> lbcurves = new List<LyrebirdCurve>();
                                            lbcurves.Add(lbc);
                                            
                                            RevitObject ro = new RevitObject();
                                            ro.Curves = lbcurves;
                                            ro.FamilyName = familyName;
                                            ro.Category = category;
                                            ro.TypeName = typeName;
                                            ro.Origin = null;
                                            ro.GHPath = curves.Paths[i].ToString();
                                            tempObjs.Add(ro);
                                        }
                                    }
                                    obj = tempObjs;
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
                                    string[] pi = paramInfo.Split(new string[] { "\n", ":" }, StringSplitOptions.None);
                                    string paramName = null;
                                    string paramStorageType = null;
                                    try
                                    {
                                        paramName = pi[1].Substring(1);
                                        paramStorageType = pi[5].Substring(1);
                                        rp.ParameterName = paramName;
                                        rp.StorageType = paramStorageType;
                                    }
                                    catch { }
                                    if (paramName != null)
                                    {
                                        GH_Structure<IGH_Goo> data = null;
                                        try
                                        {
                                            DA.GetDataTree(i, out data);
                                        }
                                        catch { }
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
                            bool send = channel.CreateOrModify(obj, this.InstanceGuid);
                        }
                        channel.Dispose();
                        try
                        {
                        }
                        catch { }
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Error\n" + "The Lyrebird Service could not be found.  Ensure Revit is running, the Lyrebird server plugin is installed, and the server is active.");
                    }
                }
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
                    if (rp.IsType == true)
                    {
                        type = "Type";
                    }
                    sb.AppendLine(string.Format("Parameter{0}: {1}  /  {2}  /  {3}", (i + 1).ToString(), rp.ParameterName, rp.StorageType, type));
                }
                message = sb.ToString();
            }
            else
            {
                message = "No data type set.  Double-click to set data type";
            }

            List<string> oids = new List<string>();
            foreach (LyrebirdId id in uniqueIDs)
            {
                oids.Add(id.UniqueId);
            }

            DA.SetData(0, message);
            DA.SetData(1, this.InstanceGuid.ToString());
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
                        this.ExpireSolution(true);
                        SyncInputs();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("The Lyrebird Service could not be found.  Ensure Revit is running, the Lyrebird server plugin is installed, and the server is active.");
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
            //Menu_AppendItem(iMenu, "Test", Menu_TestClicked);
            Menu_AppendItem(iMenu, "Full Parameter Names", Menu_ParamNamesClicked, true, paramNamesEnabled);
        }

        private void Menu_ParamNamesClicked(object sender, EventArgs e)
        {
            if (paramNamesEnabled)
            {
                paramNamesEnabled = false;
            }
            else
            {
                paramNamesEnabled = true;
                //RefreshParameters();
            }
        }

        // Simple test to see if the communication is working  Just gets the project document name
        private void Menu_TestClicked(object sender, EventArgs e)
        {
            // Currently removed from the menu
            try
            {
                LyrebirdChannel channel = new LyrebirdChannel(appVersion);
                channel.Create();

                if (channel != null)
                {
                    StringBuilder sb = new StringBuilder();
                    string test = channel.DocumentName();
                    if (test == null)
                        test = "Failed to COllect Document Name";
                    string write = "Document: " + test;
                    System.Windows.Forms.MessageBox.Show(write);
                }
                channel.Dispose();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // Originally this was done from the right-click menu.  It's since been changed 
        private void SyncInputs()
        {
            if (inputParameters.Count > 0)
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
                        Grasshopper.Kernel.Parameters.Param_GenericObject param = new Grasshopper.Kernel.Parameters.Param_GenericObject();
                        param.Name = "Parameter" + (i - 6).ToString();
                        param.NickName = "P" + (i - 6).ToString();
                        param.Description = "Parameter Name: " + inputParameters[i - 7].ParameterName + "\nIs Type: " + inputParameters[i - 7].IsType.ToString() + "\nStorageType: " + inputParameters[i - 7].StorageType;
                        param.Optional = true;
                        param.Access = GH_ParamAccess.tree;
                        Params.RegisterInputParam(param);
                    }
                }

                // Remove unnecessay inputs
                else if (inputParameters.Count < Params.Input.Count - 6)
                {
                    //System.Windows.Forms.MessageBox.Show("Going to try and remove parameters.");
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
                            param.NickName = "P" + (i + 1).ToString();
                        }
                    }
                    param.Name = "Parameter" + (i + 1).ToString();
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
            for (int i = 0; i < inputParameters.Count; i++)
            {
                try
                {
                    RevitParameter rp = inputParameters[i];
                    writer.SetString("ParameterName" + i.ToString(), rp.ParameterName);
                    writer.SetString("StorageType" + i.ToString(), rp.StorageType);
                    writer.SetBoolean("IsType" + i.ToString(), rp.IsType);
                }
                catch { }
            }
            return base.Write(writer);
        }

        public override bool Read(GH_IO.Serialization.GH_IReader reader)
        {
            FamilyName = reader.GetString("FamilyName");
            TypeName = reader.GetString("TypeName");
            Category = reader.GetString("Category");
            bool test = true;
            int i = 0;
            List<RevitParameter> parameters = new List<RevitParameter>();
            while (test)
            {
                RevitParameter rp = new RevitParameter();
                try
                {
                    rp.ParameterName = reader.GetString("ParameterName" + i.ToString());
                    rp.StorageType = reader.GetString("StorageType" + i.ToString());
                    rp.IsType = reader.GetBoolean("IsType" + i.ToString());
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
                    Rhino.Geometry.Vector3d v = orientations[i][0].Value;

                    LyrebirdPoint p = new LyrebirdPoint();
                    p.X = v.X;
                    p.Y = v.Y;
                    p.Z = v.Z;
                    ro.Orientation = p;
                    tempObj.Add(ro);
                }
                catch { }
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
                    Rhino.Geometry.Vector3d v = orientations[i][0].Value;

                    LyrebirdPoint p = new LyrebirdPoint();
                    p.X = v.X;
                    p.Y = v.Y;
                    p.Z = v.Z;
                    ro.FaceOrientation = p;
                    tempObj.Add(ro);
                }
                catch { }
            }
            return tempObj;
        }

        private LyrebirdCurve GetLBCurve(Rhino.Geometry.Curve crv)
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
            else if (crv.IsArc())
            {
                // standard arc element
                points.Add(new LyrebirdPoint(crv.PointAtStart.X, crv.PointAtStart.Y, crv.PointAtStart.Z));
                points.Add(new LyrebirdPoint(crv.PointAt(0.5).X, crv.PointAt(0.5).Y, crv.PointAt(0.5).Z));
                points.Add(new LyrebirdPoint(crv.PointAtEnd.X, crv.PointAtEnd.Y, crv.PointAtEnd.Z));
                lbc = new LyrebirdCurve(points, "Arc");
            }
            else
            {
                // Spline
                if (crv.Degree >= 3)
                {
                    Rhino.Geometry.NurbsCurve nc = crv as Rhino.Geometry.NurbsCurve;
                    if (nc != null)
                    {
                        List<LyrebirdPoint> lbPoints = new List<LyrebirdPoint>();
                        List<double> weights = new List<double>();
                        List<double> knots = new List<double>();

                        foreach (Rhino.Geometry.ControlPoint cp in nc.Points)
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
                        lbc = new LyrebirdCurve(lbPoints, weights, knots, nc.Degree, nc.IsPeriodic);
                        lbc.CurveType = "Spline";
                    }
                }
                else
                {
                    double incr = 1.0 / 100;
                    List<LyrebirdPoint> pts = new List<LyrebirdPoint>();
                    List<double> weights = new List<double>();
                    for (int i = 0; i <= 100; i++)
                    {
                        Rhino.Geometry.Point3d pt = crv.PointAtNormalizedLength(i * incr);
                        LyrebirdPoint lbp = new LyrebirdPoint(pt.X, pt.Y, pt.Z);
                        weights.Add(1.0);
                        pts.Add(lbp);
                    }

                    lbc = new LyrebirdCurve(pts, "Spline");
                    lbc.Weights = weights;
                    lbc.Degree = crv.Degree;
                }
            }

            return lbc;
        }

        protected bool CurveSegments(List<Rhino.Geometry.Curve> L, Rhino.Geometry.Curve crv, bool recursive)
        {
            if (crv == null) { return false; }
            Rhino.Geometry.PolyCurve polycurve = crv as Rhino.Geometry.PolyCurve;
            if (polycurve != null)
            {
                if (recursive) { polycurve.RemoveNesting(); }
                Rhino.Geometry.Curve[] segments = polycurve.Explode();
                if (segments == null) { return false; }
                if (segments.Length == 0) { return false; }
                if (recursive)
                {
                    foreach (Rhino.Geometry.Curve S in segments)
                    {
                        CurveSegments(L, S, recursive);
                    }
                }
                else
                {
                    foreach (Rhino.Geometry.Curve S in segments)
                    {
                        L.Add(S.DuplicateShallow() as Rhino.Geometry.Curve);
                    }
                }
                return true;
            }
            Rhino.Geometry.PolylineCurve polyline = crv as Rhino.Geometry.PolylineCurve;
            if (polyline != null)
            {
                if (recursive)
                {
                    for (int i = 0; i < (polyline.PointCount - 1); i++)
                    {
                        L.Add(new Rhino.Geometry.LineCurve(polyline.Point(i), polyline.Point(i + 1)));
                    }
                }
                else
                {
                    L.Add(polyline.DuplicateCurve());
                }
                return true;
            }
            Rhino.Geometry.Polyline p;
            if (crv.TryGetPolyline(out p))
            {
                if (recursive)
                {
                    for (int i = 0; i < (p.Count - 1); i++)
                    {
                        L.Add(new Rhino.Geometry.LineCurve(p[i], p[i + 1]));
                    }
                }
                else
                {
                    L.Add(new Rhino.Geometry.PolylineCurve(p));
                }
                return true;
            }
            //Maybe it's a LineCurve?
            Rhino.Geometry.LineCurve line = crv as Rhino.Geometry.LineCurve;
            if (line != null)
            {
                L.Add(line.DuplicateCurve());
                return true;
            }
            //It might still be an ArcCurve...
            Rhino.Geometry.ArcCurve arc = crv as Rhino.Geometry.ArcCurve;
            if (arc != null)
            {
                L.Add(arc.DuplicateCurve());
                return true;
            }
            //Nothing else worked, lets assume it's a nurbs curve and go from there...
            Rhino.Geometry.NurbsCurve nurbs = crv.ToNurbsCurve();
            if (nurbs == null) { return false; }
            double t0 = nurbs.Domain.Min;
            double t1 = nurbs.Domain.Max;
            double t;
            int LN = L.Count;
            do
            {
                if (!nurbs.GetNextDiscontinuity(Rhino.Geometry.Continuity.C1_locus_continuous, t0, t1, out t)) { break; }
                Rhino.Geometry.Interval trim = new Rhino.Geometry.Interval(t0, t);
                if (trim.Length < 1e-10)
                {
                    t0 = t;
                    continue;
                }
                Rhino.Geometry.Curve M = nurbs.DuplicateCurve();
                M = M.Trim(trim);
                if (M.IsValid) { L.Add(M); }
                t0 = t;
            } while (true);
            if (L.Count == LN) { L.Add(nurbs); }
            return true;
        }
    }
}
