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
            pManager.AddCurveParameter("Curve", "C", "Arc or line like cuves.  Closed profile curves can be used to generate floor, wall or roof sketches, or single segment non-closed curves can be used for line based family generation.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation", "O", "Vector to orient objects.", GH_ParamAccess.tree);
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
            pManager[5].Optional = true;
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
            // TODO: Establish a method of tracking previously created elements in GH or Revit.
            // Returning ID's from Revit doesn't work since it runs on a separate thread and this continues before 
            pManager.AddTextParameter("IDs", "IDs", "IDs from GH", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool runCommand = false;
            GH_Structure<GH_Point> origPoints = new GH_Structure<GH_Point>();
            GH_Structure<GH_Vector> orientations = new GH_Structure<GH_Vector>();
            DA.GetData(0, ref runCommand);
            DA.GetData(1, ref appVersion);
            DA.GetDataTree(2, out origPoints);
            DA.GetDataTree(5, out orientations);
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
                        if (origPoints != null && origPoints.Branches.Count > 0)
                        {
                            List<RevitObject> tempObjs = new List<RevitObject>();

                            for (int i = 0; i < origPoints.Branches.Count; i++)
                            {
                                GH_Point ghpt = origPoints[i][0];
                                Point point = new Point();
                                point.X = ghpt.Value.X;
                                point.Y = ghpt.Value.Y;
                                point.Z = ghpt.Value.Z;

                                RevitObject ro = new RevitObject();
                                ro.Origin = point;
                                ro.FamilyName = familyName;
                                ro.TypeName = typeName;
                                ro.GHPath = origPoints.Paths[i].ToString();
                                tempObjs.Add(ro);
                            }
                            obj = tempObjs;
                        }
                        // TODO: All other options
                        // else if adaptivePoints....
                        // else if curves...

                        // Orientation
                        if (orientations != null && orientations.Branches.Count > 0)
                        {
                            List<RevitObject> tempList = AssignOrientation(obj, orientations);
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

                        }
                        bool send = channel.CreateOrModify(obj, this.InstanceGuid);
                        
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
            //System.Windows.Forms.MessageBox.Show("Endpoint:\n" + client.Endpoint.Address.ToString());
            if (channel != null)
            {
                try
                {
                    SetRevitDataForm form = new SetRevitDataForm(channel, this);
                    form.ShowDialog();
                    if (form.DialogResult.HasValue && form.DialogResult.Value)
                    {
                        //testString = "The form has been successfully opened.";
                        this.ExpireSolution(true);
                        SyncInputs();
                    }
                }
                catch (Exception ex)
                {
                    //System.Windows.Forms.MessageBox.Show("Error\n" + ex.ToString());
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
            try
            {
                LyrebirdChannel channel = new LyrebirdChannel(appVersion);
                channel.Create();

                if (channel != null)
                {
                    StringBuilder sb = new StringBuilder();
                    //sb.AppendLine("Client State: " + client.State.ToString());
                    //sb.AppendLine("Endpoint: " + client.Endpoint.Address.ToString());
                    //string temp = client.GetDocumentName();
                    //if (temp == null)
                    //    temp = "NULL";
                    string test = channel.DocumentName();
                    if (test == null)
                        test = "FUNKY";
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

                    Point p = new Point();
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
    }
}
