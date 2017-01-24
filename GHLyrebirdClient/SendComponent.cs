using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Microsoft.Win32;

namespace Lyrebird
{
    public class SendComponent : GH_Component
    {
        string serverVersion = "Revit2016";
        bool reset = true;
        List<string> appVersions;

        System.Windows.Forms.ToolStripMenuItem appItem;

        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public SendComponent()
          : base("Lyrebird Client", "LBClient",
              "Connect to the Lyrebird Server in Revit 2017",
              "Misc", "Elk")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Send", "S", "Send the stream of data from Grasshopper to " + serverVersion, GH_ParamAccess.item, false);
            pManager.AddPointParameter("Origin Point", "O", "Origin point for sent objects.  Do not use if creating line based or adaptive families.", GH_ParamAccess.item, null);
            pManager.AddPointParameter("Adaptive Points", "A", "Points to instantiate an adaptive component.", GH_ParamAccess.tree, null);
            pManager.AddCurveParameter("Location Curve", "C", "Curves to instantiate sent objects.  Use curves for things like walls, floors, or other line-based families", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Orientation Vector", "V", "Vector to orient objects.  This will be used to determine which direction \"Front\" is with the created family or the normal vector for face based families.", GH_ParamAccess.tree);
            pManager.AddVectorParameter("Face Vector", "F", "Add an optional vector perpendicular to the face normal to control which direction the \"Top\" is on a face based family", GH_ParamAccess.tree);
            pManager[1].Optional = true;
            pManager[2].Optional = true;
            pManager[3].Optional = true;
            pManager[4].Optional = true;
            pManager[5].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Selected Family-Type Information", "I", "Information of the type of family Lyrebird will instantiate when sent.", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "M", "Message indicating successful sending of information or problems if found.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool send = false;
            Point3d pt = Point3d.Origin;
            DA.GetData(0, ref send);
            DA.GetData(1, ref pt);

            if (send && reset)
            {
                // set reset to false.  This will prevent lyrebird from sending information if you forget to turn the send trigger off.
                reset = false;

                // Create the Channel
                LBChannel channel = new LBChannel(serverVersion);

                if (channel.Create())
                {
                    string docName = channel.GetDocumentName();
                    DA.SetData(1, docName);
                }
                else
                    DA.SetData(1, "Did not successfully create a channel");
            }
            else if (!send)
                reset = true;
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
            get { return new Guid("d56bf8a8-1b2c-4ef0-ba07-53fb2c8099ef"); }
        }


        protected override void AppendAdditionalComponentMenuItems(System.Windows.Forms.ToolStripDropDown iMenu)
        {
            Menu_AppendItem(iMenu, "Full Parameter Names", Menu_ParamNamesClicked, true, false);
            Menu_AppendItem(iMenu, "UI Integration", Menu_UIClicked);

            // Try to look in the registry to see which versions of Revit are installed
            // Get current year+2 in order to get current version (+1) and beta (+2)
            int maxYear = DateTime.Now.Year + 2;
            appItem = Menu_AppendItem(iMenu, "Application");
            appVersions = new List<string>();
            bool foundVersion = false;
            for(int i = 0; i < 5; i++)
            {
                // Look for a registry key for the year
                int currentYear = maxYear - i;
                RegistryKey regKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\Revit\Autodesk Revit " + currentYear.ToString() + "\\Components");
                if (regKey != null)
                {
                    appVersions.Add("Revit" + currentYear);
                    if (serverVersion.Contains(currentYear.ToString()))
                        foundVersion = true;
                }
            }

            // Add the menu items
            for(int i = 0; i < appVersions.Count; i++)
            {
                string version = appVersions[i];
                if (!foundVersion && i == 0)
                {
                    appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, true));
                    serverVersion = version;
                }
                else if (version == serverVersion)
                    appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, true));
                else
                    appItem.DropDownItems.Add(Menu_AppendItem(iMenu, version, Menu_VersionClicked, true, false));
            }
        }

        private void Menu_ParamNamesClicked(object sender, EventArgs e)
        {

        }

        private void Menu_UIClicked(object sender, EventArgs e)
        {

        }

        private void Menu_VersionClicked(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolStripMenuItem item = sender as System.Windows.Forms.ToolStripMenuItem;
            serverVersion = item.Text;
        }
    }
}
