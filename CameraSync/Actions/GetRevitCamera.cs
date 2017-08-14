using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    class GetRevitCamera : ILyrebirdAction
    {
        public GetRevitCamera(string installPath) : base(installPath)
        {
            try
            {
                Assembly.LoadFrom(installPath + "\\RevitAPI.dll");
                Assembly.LoadFrom(installPath + "\\RevitAPIUI.dll");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error:\n" + e.Message);
            }
        }


        public override bool Command(UIApplication uiApp, Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            try
            {
                Autodesk.Revit.DB.View activeView = uiApp.ActiveUIDocument.Document.ActiveView;
                if (activeView.ViewType != ViewType.ThreeD)
                {
                    TaskDialog.Show("Warning",
                        "Attempted to use Lyrebird to access the Revit Camera from a non-3d View. Make sure you change your active view to a 3D view before running this command.");
                    outputs = null;
                    return false;
                }

                View3D view3d = activeView as View3D;
                bool isPerspective = view3d.IsPerspective;
                XYZ cLoc = view3d.GetOrientation().EyePosition;
                XYZ fDir = view3d.GetOrientation().ForwardDirection;
                XYZ uDir = view3d.GetOrientation().UpDirection;

                //System.Windows.Forms.MessageBox.Show("CommandTest" +
                //                                     "\nviewName:\t" + view3d.ViewName +
                //                                     "\nisPerspective:\t" + isPerspective.ToString() +
                //                                     "\ncameraPos:\t" + cLoc.ToString() +
                //                                     "\ncamFwdDir:\t" + fDir.ToString() +
                //                                     "\ncamUpDir:\t" + uDir.ToString());
                
                LINE.Geometry.Point3d cameraLocation =
                    LINE.Geometry.RevitConvert.Point3dFromRevit(cLoc);
                LINE.Geometry.Vector3d cameraDir =
                    LINE.Geometry.RevitConvert.Vector3dFromRevit(fDir);
                LINE.Geometry.Vector3d cameraUpDir =
                    LINE.Geometry.RevitConvert.Vector3dFromRevit(uDir);


                //System.Windows.Forms.MessageBox.Show("Converted:" +
                //                                     "\nviewName:\t" + view3d.ViewName +
                //                                     "\nisPerspective:\t" + isPerspective.ToString() +
                //                                     "\ncameraPos:\t" + cameraLocation.ToString() +
                //                                     "\ncamFwdDir:\t" + cameraDir.ToString() +
                //                                     "\ncamUpDir:\t" + cameraUpDir.ToString());

                outputs = new Dictionary<string, object>
                {
                    {"viewName", view3d.ViewName},
                    {"isPerspective", isPerspective},
                    {"cameraLoc", JsonConvert.SerializeObject(cameraLocation, Formatting.Indented)},
                    {"cameraDir", JsonConvert.SerializeObject(cameraDir, Formatting.Indented)},
                    {"cameraUp", JsonConvert.SerializeObject(cameraUpDir, Formatting.Indented)}
                };
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error:\n" + ex.Message);
                outputs = null;
                return false;
            }
        }

        /// <summary>
        /// This property is not part of the ILyrebirdAction abstract class, but is vital for proper functioning. This
        /// is manually synced to have the same Guid as the GHComponent this Action is paired with.
        /// </summary>
        public static Guid CommandGuid => new Guid("f8ef435b-4a41-406b-9506-b10c8f0490fa");
    }
}
