using System;
using System.Collections.Generic;

using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

using LMNA.Lyrebird.LyrebirdCommon;
using System.Diagnostics;


namespace LMNA.Lyrebird.GH
{
    public class RevitCategoryElemComp : GH_Component
    {

        int appVersion = Properties.Settings.Default.RevitVersion;

        ElementIdCategory eic = ElementIdCategory.Material;

        bool r2014 = true;
        bool r2015 = false;
        bool r2016 = false;

        List<string> elements;

        public RevitCategoryElemComp()
            : base("Revit Category Elements", "RvtElem", "Revit category elements for ElementId parameters", "LMNts", "Utilities")
        {
            if (appVersion == 1)
            {
                r2014 = true;
                r2015 = false;
                r2016 = false;
            }
            else if (appVersion == 2)
            {
                r2014 = false;
                r2015 = true;
                r2016 = false;
            }
            else if (appVersion == 3)
            {
                r2014 = false;
                r2015 = false;
                r2016 = true;
            }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Trigger the retrieval of Revit materials", GH_ParamAccess.item, false);
            pManager.AddIntegerParameter("Category", "C", "n1 = Images\n2 = Levels\n3 = Materials\n4 = Phases", GH_ParamAccess.item, 3);
            Grasshopper.Kernel.Parameters.Param_Integer paramInt = pManager[1] as Grasshopper.Kernel.Parameters.Param_Integer;
            if (paramInt != null)
            {
                //paramInt.AddNamedValue("Design Options", 0);
                paramInt.AddNamedValue("Images", 1);
                paramInt.AddNamedValue("Levels", 2);
                paramInt.AddNamedValue("Materials", 3);
                paramInt.AddNamedValue("Phases", 4);
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Elements", "E", "Elements retrieved from the Revit document.\nThey are in the format of NAME,ELEMENTID", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool runCommand = false;
            int category = 3;
            DA.GetData(0, ref runCommand);
            DA.GetData(1, ref category);

            try
            {
                ElementIdCategory eicTemp = (ElementIdCategory)category;
                if (eic != eicTemp)
                {
                    eic = (ElementIdCategory)category;
                    elements.Clear();
                }
            }
            catch { }
            
            if (runCommand)
            {
                // Open the Channel to Revit
                LyrebirdChannel channel = new LyrebirdChannel(appVersion);
                channel.Create();

                if (channel != null)
                {
                    string documentName = channel.DocumentName();
                    if (documentName != null)
                    {
                        elements = channel.GetCategoryElements(eic);
                    }

                    channel.Dispose();
                }
            }
            DA.SetDataList(0, elements);
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
            get { return new Guid("{fcb68ef3-27cc-4883-bb2b-2a29c52e33da}"); }
        }

        protected override void AppendAdditionalComponentMenuItems(System.Windows.Forms.ToolStripDropDown iMenu)
        {
            // Revit Version Selector
            System.Windows.Forms.ToolStripMenuItem appItem = Menu_AppendItem(iMenu, "Application");
            appItem.DropDownItems.Add(Menu_AppendItem(iMenu, "Revit 2014", Menu_R2014Clicked, true, r2014));
            appItem.DropDownItems.Add(Menu_AppendItem(iMenu, "Revit 2015", Menu_R2015Clicked, true, r2015));
            appItem.DropDownItems.Add(Menu_AppendItem(iMenu, "Revit 2016", Menu_R2016Clicked, true, r2016));
        }

        private void Menu_R2014Clicked(object sender, EventArgs e)
        {
            r2014 = !r2014;
            if (r2014)
            {
                r2015 = false;
                r2016 = false;
            }
            appVersion = 1;
            Properties.Settings.Default.RevitVersion = appVersion;
            Properties.Settings.Default.Save();
        }

        private void Menu_R2015Clicked(object sender, EventArgs e)
        {
            r2015 = !r2015;
            if (r2015)
            {
                r2014 = false;
                r2016 = false;
            }
            appVersion = 2;
            Properties.Settings.Default.RevitVersion = appVersion;
            Properties.Settings.Default.Save();
        }

        private void Menu_R2016Clicked(object sender, EventArgs e)
        {
            r2016 = !r2016;
            if (r2016)
            {
                r2014 = false;
                r2015 = false;
            }
            appVersion = 3;
            Properties.Settings.Default.RevitVersion = appVersion;
            Properties.Settings.Default.Save();
        }
    }
}