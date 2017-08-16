using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using RG = Rhino.Geometry;

using LG = LINE.Geometry;

namespace Lyrebird
{
    public class GetCameraComponent : GH_Component
    {
        private bool _reset = true;
        private string _serverVersion = "Revit2017";

        private string viewName = null;
        RG.Point3d cameraPosition = RG.Point3d.Unset;
        RG.Point3d cameraTarget = RG.Point3d.Unset;
        private bool isPerspective = false;
        RG.Vector3d cameraUpVector = RG.Vector3d.Unset;

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public GetCameraComponent()
          : base("GetRevitCamera", "GetRevCam",
              "Get the Revit Camera properties to sync with a Rhino Camera.",
              "Misc", "Lyrebird")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Run LyrebirdAction", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("View Name", "N", "Revit view name", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Is Perspective", "P", "Is the camera a Perspective", GH_ParamAccess.item);
            pManager.AddPointParameter("Camera Location", "L", "Location of the camera", GH_ParamAccess.item);
            pManager.AddPointParameter("Camera Target", "T", "Camera target position", GH_ParamAccess.item);
            pManager.AddVectorParameter("Up Direction", "U", "Camera Up Direction", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var send = false;
            DA.GetData(0, ref send);

            if (send && _reset)
            {
                // set _reset to false.  This will prevent lyrebird from sending information if you forget to turn the send trigger off.
                _reset = false;

                // Create the Channel
                var channel = new LBChannel(_serverVersion);

                if (channel.Create())
                {
                    Dictionary<string, object> input = new Dictionary<string, object> { { "CommandGuid", ComponentGuid } };
                    Dictionary<string, object> output = channel.LBAction(input);
                    
                    //System.Windows.Forms.MessageBox.Show("output Null? " + (output == null).ToString());
                    if (output == null || !output.ContainsKey("viewName"))
                        return;

                    // TODO: Should probably have this set the RhinoViewport camera
                    //       No reason to rely on another plugin (Horester probably) to set the camera when
                    //       this plugin can manage all of that on its own, since that's what it's designed to do.
                    //       Consider whether it should do it when triggered (pull info, set camera) or as a second
                    //       trigger to set camera to retrieved settings.
                    try
                    {
                        viewName = output["viewName"].ToString();
                        LG.Point3d camPos = LG.JsonConvert.Point3dFromJson(output["cameraLoc"].ToString());
                        cameraPosition = LG.RhinoConvert.Point3dToRhino(camPos);
                        LG.Vector3d camDir = LG.JsonConvert.Vector3dFromJson(output["cameraDir"].ToString());
                        var dir = LG.RhinoConvert.Vector3dToRhino(camDir);
                        cameraTarget = cameraPosition + dir;
                        isPerspective = (bool)output["isPerspective"];
                        LG.Vector3d camUp = LG.JsonConvert.Vector3dFromJson(output["cameraUp"].ToString());
                        cameraUpVector = LG.RhinoConvert.Vector3dToRhino(camUp);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show("Error\n" + e.Message);
                        return;
                    }
                }
                channel.Dispose();
            }
            else if (!send)
            {
                _reset = true;
            }

            DA.SetData(0, viewName);
            DA.SetData(1, isPerspective);
            DA.SetData(2, cameraPosition);
            DA.SetData(3, cameraTarget);
            DA.SetData(4, cameraUpVector);
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                //return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("f8ef435b-4a41-406b-9506-b10c8f0490fa"); }
        }
    }
}
