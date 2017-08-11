using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace Lyrebird
{
    public class ActionTestComponent : GH_Component
    {
        private bool _reset = true;
        private string _serverVersion = "Revit2017";

        /// <summary>
        /// Initializes a new instance of the ActionTestComponent class.
        /// </summary>
        public ActionTestComponent()
          : base("ActionTestComponent", "ACTION!!",
              "Test out a LB Action, running externally from Revit.",
              "Misc", "Elk")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Trigger", "T", "Run LyrebirdAction", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Document", "D", "Name of the current Revit document", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
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
                    string assemblyPath = this.GetType().Assembly.Location;
                    if (string.IsNullOrEmpty(assemblyPath))
                        return;

                    // Get the RevitAPI path
                    List<string> revitApiPath = channel.GetRevitAPIPath();
                    System.Windows.Forms.MessageBox.Show("revitApiPath is null? " + (revitApiPath == null).ToString());
                    if (revitApiPath == null)
                        return;

                    foreach (string refPath in revitApiPath)
                    {
                        AppDomain.CurrentDomain.Load(refPath);
                    }

                    Dictionary<string, object> input = new Dictionary<string, object> {{"CommandGuid", Lyrebird.GetRevitDocName.CommandGuid}, {"AssemblyPath", assemblyPath } };
                    var output = channel.LBAction(input);
                    if (output == null || !output.ContainsKey("docName"))
                    {
                        DA.SetData(0, "output is null or doesn't have 'docName'");
                        return;
                    }

                    var docName = output["docName"];
                    DA.SetData(0, docName);
                }
                else
                {
                    DA.SetData(0, "Did not successfully create a channel");
                }
            }
            else if (!send)
            {
                _reset = true;
            }
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => null;

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("8a9ce1a6-9861-49f0-8b3b-535779d197b5");
    }
}