using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace LMNA.Lyrebird.LyrebirdCommon
{
    [DataContract]
    public class RevitParameter
    {

        [DataMember]
        public string ParameterName { get; set; }

        [DataMember]
        public string StorageType { get; set; }

        [DataMember]
        public string Value { get; set; }

        [DataMember]
        public bool IsType { get; set; }

        public RevitParameter()
        {
            ParameterName = null;
            StorageType = null;
            Value = null;
            IsType = true;
        }

        public RevitParameter(string parameterName, string storageType, string parameterValue, bool isType)
        {
            ParameterName = parameterName;
            StorageType = storageType;
            Value = parameterValue;
            IsType = isType;
        }
    }

    [DataContract]
    public class RevitObject
    {
        [DataMember]
        public string Category { get; set; }

        [DataMember]
        public string FamilyName { get; set; }

        [DataMember]
        public string TypeName { get; set; }

        [DataMember]
        public List<RevitParameter> Parameters { get; set; }

        [DataMember]
        public Point Origin { get; set; }

        [DataMember]
        public List<Point> AdaptivePoints { get; set; }

        [DataMember]
        public LyrebirdId UniqueID { get; set; }

        [DataMember]
        public List<List<Point>> CurvePoints { get; set; }

        [DataMember]
        public string GHPath { get; set; }

        //public Autodesk.Revit.DB.Curve Curve
        //{
        //    get { return curve; }
        //    set { curve = value; }
        //}

        [DataMember]
        public Point Orientation { get; set; }

        public RevitObject(string family, string type, List<RevitParameter> parameters, Point origin, Point orient)
        {
            FamilyName = family;
            TypeName = type;
            Parameters = parameters;
            Origin = origin;
            Orientation = orient;
        }

        public RevitObject(string category, string family)
        {
            FamilyName = family;
            Category = category;
        }

        public RevitObject()
        {
            FamilyName = null;
            TypeName = null;
            Parameters = new List<RevitParameter>();
            Origin = Point.Zero;
            Orientation = Point.Zero;
        }

        public RevitObject(string family, string type, List<RevitParameter> parameters, List<Point> adaptivePoints)
        {
            FamilyName = family;
            TypeName = type;
            Parameters = parameters;
            AdaptivePoints = adaptivePoints;
        }
    }

    // TODO: Originally made to pass revit unique id's to GH, but since it's
    // running on a separate thread it doesn't get returned in sequence so
    // get's left behind.  Evaluate a different way of handling it.
    [DataContract]
    public class LyrebirdId
    {
        [DataMember]
        public string UniqueId { get; set; }

        [DataMember]
        public string GHPath { get; set; }

        public LyrebirdId(string uniqueId, string ghPath)
        {
            UniqueId = uniqueId;
            GHPath = ghPath;
        }
    }

    [DataContract]
    public class Point
    {
        [DataMember]
        public double X { get; set; }

        [DataMember]
        public double Y { get; set; }

        [DataMember]
        public double Z { get; set; }

        public Point()
        {
            X = 0.0;
            Y = 0.0;
            Z = 0.0;
        }

        public Point(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public static Point Zero
        {
            get { return new Point(0.0, 0.0, 0.0); }
        }
    }

    [DataContract]
    public class LyrebirdCurve
    {
        [DataMember]
        public List<Point> Points { get; set; }

        [DataMember]
        public int Degree { get; set; }

        [DataMember]
        public string CurveType { get; set; }

        public LyrebirdCurve(List<Point> points, string curveType)
        {
            Points = points;
            CurveType = curveType;
        }

        public LyrebirdCurve(List<Point> points, string curveType, int degree)
        {
            Points = points;
            CurveType = curveType;
            Degree = degree;
        }

    }
}
