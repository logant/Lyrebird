using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Grasshopper.Kernel;
using Microsoft.Win32;
using RG = Rhino.Geometry;

using LG = LINE.Geometry;

namespace Lyrebird
{
    public class GetCameraComponent : GH_Component
    {
        private ToolStripMenuItem _appItem;
        private List<string> _appVersions;
        private bool _reset = true;
        private string _serverVersion = "Revit2017";

        private string _viewName = null;
        private RG.Point3d _cameraPosition = RG.Point3d.Unset;
        private RG.Point3d _cameraTarget = RG.Point3d.Unset;
        private bool _isPerspective = false;
        private RG.Vector3d _cameraUpVector = RG.Vector3d.Unset;

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
                        _viewName = output["viewName"].ToString();
                        LG.Point3d camPos = LG.JsonConvert.Point3dFromJson(output["cameraLoc"].ToString());
                        _cameraPosition = LG.RhinoConvert.Point3dToRhino(camPos);
                        LG.Vector3d camDir = LG.JsonConvert.Vector3dFromJson(output["cameraDir"].ToString());
                        var dir = LG.RhinoConvert.Vector3dToRhino(camDir);
                        _cameraTarget = _cameraPosition + dir;
                        _isPerspective = (bool)output["isPerspective"];
                        LG.Vector3d camUp = LG.JsonConvert.Vector3dFromJson(output["cameraUp"].ToString());
                        _cameraUpVector = LG.RhinoConvert.Vector3dToRhino(camUp);
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

            DA.SetData(0, _viewName);
            DA.SetData(1, _isPerspective);
            DA.SetData(2, _cameraPosition);
            DA.SetData(3, _cameraTarget);
            DA.SetData(4, _cameraUpVector);
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        // You can add image files to your project resources and access them like this:
        //return Resources.IconForThisComponent;
        protected override System.Drawing.Bitmap Icon => null;

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("f8ef435b-4a41-406b-9506-b10c8f0490fa");

        /// <summary>
        /// Primarily this is adding a menu item for each version of Revit that is found on the machine.
        /// It does this by looking through the Registry for Revit. You can optionally target specific versions
        /// of Revit explicitly instead of searching in case you don't want to support 
        /// all versions with this particular component.
        /// </summary>
        /// <param name="iMenu"></param>
        public override void AppendAdditionalMenuItems(ToolStripDropDown iMenu)
        {
            // Try to look in the registry to see which versions of Revit are installed
            // Get current year+2 in order to get current version (+1) and beta (+2)
            var maxYear = DateTime.Now.Year + 2;
            _appItem = Menu_AppendItem(iMenu, "Application");
            _appVersions = new List<string>();
            var foundVersion = false;
            for (var i = 0; i < 5; i++)
            {
                // Look for a registry key for the year
                var currentYear = maxYear - i;
                var regKey =
                    Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\Revit\Autodesk Revit " + currentYear +
                                                     "\\Components");
                if (null == regKey) continue;

                _appVersions.Add("Revit" + currentYear);
                if (_serverVersion.Contains(currentYear.ToString()))
                    foundVersion = true;
            }

            // Add the menu items
            for (var i = 0; i < _appVersions.Count; i++)
            {
                var version = _appVersions[i];
                if (!foundVersion && i == 0)
                {
                    _appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, true));
                    _serverVersion = version;
                }
                else if (version == _serverVersion)
                {
                    _appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, true));
                }
                else
                {
                    _appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, false));
                }
            }
        }

        private void Menu_VersionClicked(object sender, EventArgs e)
        {
            var item = sender as ToolStripMenuItem;
            if (item != null) _serverVersion = item.Text;
        }
    }
}
