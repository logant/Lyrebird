using System.Collections.Generic;
using System.Runtime.Serialization;

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
        public int CategoryId { get; set; }

        [DataMember]
        public string FamilyName { get; set; }

        [DataMember]
        public string TypeName { get; set; }

        [DataMember]
        public List<RevitParameter> Parameters { get; set; }

        [DataMember]
        public LyrebirdPoint Origin { get; set; }

        [DataMember]
        public List<LyrebirdPoint> AdaptivePoints { get; set; }

        [DataMember]
        public LyrebirdId UniqueID { get; set; }

        [DataMember]
        public List<LyrebirdCurve> Curves { get; set; }

        [DataMember]
        public string GHPath { get; set; }

        [DataMember]
        public string GHScaleName { get; set; }

        [DataMember]
        public double GHScaleFactor { get; set; }

        [DataMember]
        public List<int> CurveIds { get; set; }

        //public Autodesk.Revit.DB.Curve Curve
        //{
        //    get { return curve; }
        //    set { curve = value; }
        //}

        [DataMember]
        public LyrebirdPoint Orientation { get; set; }

        [DataMember]
        public LyrebirdPoint FaceOrientation { get; set; }

        public RevitObject(string family, string type, List<RevitParameter> parameters, LyrebirdPoint origin, LyrebirdPoint orient)
        {
            FamilyName = family;
            TypeName = type;
            Parameters = parameters;
            Origin = origin;
            Orientation = orient;
        }

        public RevitObject(string category,int categoryId, string family)
        {
            FamilyName = family;
            Category = category;
            CategoryId = categoryId;
        }

        public RevitObject()
        {
            FamilyName = null;
            TypeName = null;
            Parameters = new List<RevitParameter>();
            Origin = LyrebirdPoint.Zero;
            Orientation = LyrebirdPoint.Zero;
        }

        public RevitObject(string family, string type, List<RevitParameter> parameters, List<LyrebirdPoint> adaptivePoints)
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

        [DataMember]
        public string GHName { get; set; }

        public LyrebirdId(string uniqueId, string ghName, string ghPath)
        {
            UniqueId = uniqueId;
            GHName = ghName;
            GHPath = ghPath;
        }
    }

    [DataContract]
    public class LyrebirdPoint
    {
        [DataMember]
        public double X { get; set; }

        [DataMember]
        public double Y { get; set; }

        [DataMember]
        public double Z { get; set; }

        public LyrebirdPoint()
        {
            X = 0.0;
            Y = 0.0;
            Z = 0.0;
        }

        public LyrebirdPoint(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public static LyrebirdPoint Zero
        {
            get { return new LyrebirdPoint(0.0, 0.0, 0.0); }
        }
    }

    [DataContract]
    public class LyrebirdCurve
    {
        [DataMember]
        public List<LyrebirdPoint> ControlPoints { get; set; }

        [DataMember]
        public int Degree { get; set; }

        [DataMember]
        public List<double> Weights { get; set; }

        [DataMember]
        public List<double> Knots { get; set; }

        [DataMember]
        public string CurveType { get; set; }

        [DataMember]
        public bool Periodic { get; set; }

        public LyrebirdCurve()
        {

        }

        public LyrebirdCurve(List<LyrebirdPoint> points, string curveType)
        {
            ControlPoints = points;
            CurveType = curveType;
        }

        public LyrebirdCurve(List<LyrebirdPoint> points, string curveType, int degree)
        {
            ControlPoints = points;
            CurveType = curveType;
            Degree = degree;
        }

        public LyrebirdCurve(List<LyrebirdPoint> points, List<double> weights, List<double> knots, int degree, bool periodic)
        {
            ControlPoints = points;
            Weights = weights;
            Knots = knots;
            Degree = degree;
            Periodic = periodic;
        }

    }

    [DataContract]
    public class Runs
    {
        [DataMember]
        public int RunId { get; set; }

        [DataMember]
        public string RunName { get; set; }

        [DataMember]
        public string FamilyType { get; set; }

        [DataMember]
        public List<int> ElementIds { get; set; }

        public Runs(int id, string name, string familyType)
        {
            RunId = id;
            RunName = name;
            FamilyType = familyType;
        }

        public Runs()
        {
        }
    }

    [DataContract]
    public class RunCollection
    {
        [DataMember]
        public System.Guid ComponentGuid { get; set; }

        [DataMember]
        public List<Runs> Runs { get; set; }

        [DataMember]
        public string NickName { get; set; }

        public RunCollection()
        {
        }
    }
}
