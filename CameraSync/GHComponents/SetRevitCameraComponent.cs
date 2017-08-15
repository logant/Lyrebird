using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;
using LG = LINE.Geometry;
using Newtonsoft.Json;

namespace Lyrebird
{
    public class SetRevitCameraComponent : GH_Component
    {
        private bool _reset = true;
        private string _serverVersion = "Revit2017";

        /// <summary>
        /// Initializes a new instance of the SetRevitCameraComponent class.
        /// </summary>
        public SetRevitCameraComponent()
          : base("SetRevitCameraComponent", "SetRevCam",
              "Set the Revit camera to match the Rhino viewport camera.",
              "Misc", "Lyrebird")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Run LyrebirdAction", GH_ParamAccess.item, false);
            pManager.AddTextParameter("Rhino Viewport", "RhVP",
                "Name of a Rhino viewport to get the camera properties.", GH_ParamAccess.item);
            pManager.AddTextParameter("Revit View Name", "ReVP", "Name of a Revit view to set the camera.",
                GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Result", "R", "Result of the operation.", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var send = false;
            string rhinoVPName = null;
            string revitVPName = null;
            DA.GetData(0, ref send);
            DA.GetData(1, ref rhinoVPName);
            DA.GetData(2, ref revitVPName);

            if (send && _reset)
            {
                // set _reset to false.  This will prevent lyrebird from sending information if you forget to turn the send trigger off.
                _reset = false;

                // Get the Rhino Viewport information
                Rhino.Display.RhinoViewport rhinoVP = null;
                if (!string.IsNullOrEmpty(rhinoVPName))
                {
                    try
                    {
                        rhinoVP = Rhino.RhinoDoc.ActiveDoc.Views.Find(rhinoVPName, false).ActiveViewport;
                    }
                    catch
                    {
                        // Do nothing
                    }
                    
                }
                if (!string.IsNullOrEmpty(rhinoVPName))
                {
                    rhinoVP = Rhino.RhinoDoc.ActiveDoc.Views.ActiveView.ActiveViewport;
                }


                // Create the Channel
                var channel = new LBChannel(_serverVersion);

                if (channel.Create())
                {
                    LG.Point3d cameraLocation = LG.RhinoConvert.Point3dFromRhino(rhinoVP.CameraLocation);
                    LG.Vector3d cameraDirection = LG.RhinoConvert.Vector3dFromRhino(rhinoVP.CameraDirection);
                    LG.Vector3d upDirection = LG.RhinoConvert.Vector3dFromRhino(rhinoVP.CameraUp);

                    double halfDiag;
                    double halfHoriz;
                    double halfVert;
                    rhinoVP.GetCameraAngle(out halfDiag, out halfVert, out halfHoriz);

                    Dictionary<string, object> input = new Dictionary<string, object>
                    {
                        { "CommandGuid", ComponentGuid },
                        { "RevitVPName", revitVPName },
                        { "CameraLoc", JsonConvert.SerializeObject(cameraLocation, Formatting.Indented) },
                        { "CameraDir", JsonConvert.SerializeObject(cameraDirection, Formatting.Indented) },
                        { "CameraUpDir", JsonConvert.SerializeObject(upDirection, Formatting.Indented) },
                        { "HalfHoriz", halfHoriz },
                        { "HalfVert", halfVert }
                    };

                    Dictionary<string, object> output = channel.LBAction(input);

                    //System.Windows.Forms.MessageBox.Show("output Null? " + (output == null).ToString());
                    if (output == null || !output.ContainsKey("viewName"))
                        return;
                }
                channel.Dispose();
            }
            else if (!send)
            {
                _reset = true;
            }
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9e5496d1-2f48-4e46-83c8-35353eaa71ba"); }
        }
    }
}