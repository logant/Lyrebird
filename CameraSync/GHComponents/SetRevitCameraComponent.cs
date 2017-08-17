using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Grasshopper.Kernel;
using Microsoft.Win32;
using Rhino.Geometry;
using LG = LINE.Geometry;
using Newtonsoft.Json;

namespace Lyrebird
{
    public class SetRevitCameraComponent : GH_Component
    {
        private ToolStripMenuItem _appItem;
        private List<string> _appVersions;
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
                // set _reset to false. This will prevent lyrebird from sending information 
                // if you used a toggle and forget to turn the send trigger off.
                _reset = false;

                // Get the Rhino Viewport information
                Rhino.Display.RhinoViewport rhinoVP = null;
                if (!string.IsNullOrEmpty(rhinoVPName))
                {
                    try
                    {
                        rhinoVP = Rhino.RhinoDoc.ActiveDoc.Views.Find(rhinoVPName, false).ActiveViewport;
                    }
                    catch { } // Ignore
                }
                if (null == rhinoVP)
                {
                    rhinoVP = Rhino.RhinoDoc.ActiveDoc.Views.ActiveView.ActiveViewport;
                }
                if (null == revitVPName)
                    revitVPName = string.Empty;

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
        // You can add image files to your project resources and access them like this:
        // return Resources.IconForThisComponent;
        protected override System.Drawing.Bitmap Icon => null;

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("9e5496d1-2f48-4e46-83c8-35353eaa71ba");

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