using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using LG = LINE.Geometry;

namespace Lyrebird
{
    class SetRevitCamera : ILyrebirdAction
    {
        public SetRevitCamera(string installPath) : base(installPath)
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
            outputs = null;
            try
            {
                View3D revitView = null;

                string revitViewName = inputs["RevitVPName"].ToString();

                // If the view name was supplied, check to see if it can be found and make sure that it's a ViewType.ThreeD
                if(!string.IsNullOrEmpty(revitViewName))
                {
                    View3D view = new FilteredElementCollector(uiApp.ActiveUIDocument.Document).OfClass(typeof(View3D)).First(v => v.Name.ToLower() == revitViewName.ToLower()) as View3D;
                    if (view != null)
                        revitView = view;
                }
                // If the revitView is still null, check to see if the active view is a ViewType.ThreeD
                if (revitView == null && uiApp.ActiveUIDocument.Document.ActiveView.ViewType == ViewType.ThreeD)
                    revitView = uiApp.ActiveUIDocument.Document.ActiveView as View3D;

                // if the view still hasn't been set, abort
                if (revitView == null)
                    return false;

                // Get the input data for the camera and viewport boundary.
                LG.Point3d cameraPos = LG.JsonConvert.Point3dFromJson(inputs["CameraLoc"].ToString());
                LG.Vector3d cameraDir = LG.JsonConvert.Vector3dFromJson(inputs["CameraDir"].ToString());
                LG.Vector3d cameraUpDir = LG.JsonConvert.Vector3dFromJson(inputs["CameraUpDir"].ToString());
                double halfHoriz = (double) inputs["HalfHoriz"];
                double halfVert = (double)inputs["HalfVert"];

                if (cameraPos == null || cameraDir == null || cameraUpDir == null)
                    return false;

                XYZ cameraLocation = LG.RevitConvert.Point3dToRevit(cameraPos);
                XYZ cameraFwdDirection = LG.RevitConvert.Vector3dToRevit(cameraDir);
                XYZ cameraUpDirection = LG.RevitConvert.Vector3dToRevit(cameraUpDir);

                using (Transaction trans = new Transaction(uiApp.ActiveUIDocument.Document, "Lyrebird - Set Camera"))
                {
                    // TODO: Work on Implementing Units
                    //       This will pull the Revit document units, but its not working on the camera location currently
                    
                    // Get the units of the Revit doc.
                    Units units = uiApp.ActiveUIDocument.Document.GetUnits();
                    FormatOptions fOpt = units.GetFormatOptions(UnitType.UT_Length);
                    DisplayUnitType dut = fOpt.DisplayUnits;

                    trans.Start();
                    revitView.SetOrientation(new ViewOrientation3D(cameraLocation, cameraUpDirection, cameraFwdDirection));

                    BoundingBoxXYZ bbox = revitView.CropBox;
                    XYZ halfMax = bbox.Max / 2;
                    XYZ halfMin = bbox.Min / 2;
                    revitView.CropBox.Min = halfMin;
                    revitView.CropBox.Max = halfMax;

                    // Get the perspective viewport's boundary
                    ViewCropRegionShapeManager shapeManager = revitView.GetCropRegionShapeManager();
                    CurveLoop curveLoop = shapeManager.GetCropShape()[0];
                    Plane p = curveLoop.GetPlane();
                    double focalLength = UnitUtils.ConvertFromInternalUnits(p.Origin.DistanceTo(revitView.GetOrientation().EyePosition), dut);
                    
                    XYZ focalVect = p.Origin - revitView.GetOrientation().EyePosition;
                    focalVect.Normalize();
                    focalVect *= UnitUtils.ConvertToInternalUnits(50.00 - focalLength, dut);
                    Transform shiftPlane = Transform.CreateTranslation(focalVect);
                    shapeManager.GetCropShape()[0].Transform(shiftPlane);

                    // Get the points for the new boundary loop based on the plane and the halfHoriz/halfVert numbers
                    // Note 50.00 is the standard lens length in Revit. Easiest to just keep 50 as the default between
                    // Revit and Rhino for now. Perhaps more thought can be done to handle different focal lengths.
                    double halfWidth = Math.Sin(halfHoriz) * UnitUtils.ConvertToInternalUnits(50.00, dut);
                    double halfHeight = Math.Sin(halfVert) * UnitUtils.ConvertToInternalUnits(50.00, dut);

                    // TODO: The Z coordinate needs to be resolved better.
                    // I don't recall where this came from, but it really screws up the viewport boundaries at the scale
                    // I'm currently testing it. The previous project the Macro this came from was much larger and made
                    // more sense. Once that's resovled I can turn the revitView.CropBox setting back on.
                    BoundingBoxXYZ bb = new BoundingBoxXYZ();
                    XYZ LL = new XYZ(-halfWidth, -halfHeight, -2368.398325399);
                    XYZ UR = new XYZ(halfWidth, halfHeight, -2.368398325);
                    bb.Min = LL;
                    bb.Max = UR;
                    //revitView.CropBox = bb;

                    trans.Commit();
                }
                uiApp.ActiveUIDocument.RefreshActiveView();

                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error Setting Camera:\n" + e.Message);
                return false;
            }
        }

        /// <summary>
        /// This property is not part of the ILyrebirdAction abstract class, but is vital for proper functioning. This
        /// is manually synced to have the same Guid as the GHComponent this Action is paired with.
        /// </summary>
        public static Guid CommandGuid => new Guid("9e5496d1-2f48-4e46-83c8-35353eaa71ba");
    }
}
