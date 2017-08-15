//using System.Collections.Generic;
//using System.Runtime.Serialization;
//using LG = LINE.Geometry;

//namespace Lyrebird
//{
//    public enum ElementIdCategory { Invalid, DesignOption, Image, Level, Material, Phase }

//    [DataContract]
//    public class RevitParameter
//    {
//        [DataMember]
//        public string ParameterName { get; set; }

//        [DataMember]
//        public string StorageType { get; set; }

//        [DataMember]
//        public string Value { get; set; }

//        [DataMember]
//        public bool IsType { get; set; }

//        public RevitParameter()
//        {
//            ParameterName = null;
//            StorageType = null;
//            Value = null;
//            IsType = true;
//        }

//        public RevitParameter(string parameterName, string storageType, string parameterValueStr, bool isType)
//        {
//            ParameterName = parameterName;
//            StorageType = storageType;
//            Value = parameterValueStr;
//            IsType = isType;
//        }
//    }

//    [DataContract]
//    public class RevitObject
//    {
//        [DataMember]
//        public string Category { get; set; }

//        [DataMember]
//        public int CategoryId { get; set; }

//        [DataMember]
//        public string FamilyName { get; set; }

//        [DataMember]
//        public string TypeName { get; set; }

//        [DataMember]
//        public List<RevitParameter> Parameters { get; set; }

//        [DataMember]
//        public Point3d Origin { get; set; }

//        [DataMember]
//        public List<Point3d> AdaptivePoints { get; set; }

//        [DataMember]
//        public List<Curve> Curves { get; set; }

//        [DataMember]
//        public string GHPath { get; set; }

//        [DataMember]
//        public string GHScaleName { get; set; }

//        [DataMember]
//        public double GHScaleFactor { get; set; }

//        [DataMember]
//        public List<int> CurveIds { get; set; }

//        [DataMember]
//        public Point3d Orientation { get; set; }

//        [DataMember]
//        public Point3d FaceOrientation { get; set; }

//        public RevitObject(string family, string type, List<RevitParameter> parameters, Point3d origin, Point3d orient)
//        {
//            FamilyName = family;
//            TypeName = type;
//            Parameters = parameters;
//            Origin = origin;
//            Orientation = orient;
//        }

//        public RevitObject(string category, int categoryId, string family)
//        {
//            FamilyName = family;
//            Category = category;
//            CategoryId = categoryId;
//        }

//        public RevitObject()
//        {
//            FamilyName = null;
//            TypeName = null;
//            Parameters = new List<RevitParameter>();
//            Origin = LG.Point3d.Origin;
//            Orientation = LG.Point3d.Origin;
//        }

//        public RevitObject(string family, string type, List<RevitParameter> parameters, List<Point3d> adaptivePoints)
//        {
//            FamilyName = family;
//            TypeName = type;
//            Parameters = parameters;
//            AdaptivePoints = adaptivePoints;
//        }
//    }

//    [DataContract]
//    public class Runs
//    {
//        [DataMember]
//        public int RunId { get; set; }

//        [DataMember]
//        public string RunName { get; set; }

//        [DataMember]
//        public string FamilyType { get; set; }

//        [DataMember]
//        public List<int> ElementIds { get; set; }

//        public Runs(int id, string name, string familyType)
//        {
//            RunId = id;
//            RunName = name;
//            FamilyType = familyType;
//        }

//        public Runs()
//        {
//        }
//    }

//    [DataContract]
//    public class RunCollection
//    {
//        [DataMember]
//        public System.Guid ComponentGuid { get; set; }

//        [DataMember]
//        public List<Runs> Runs { get; set; }

//        [DataMember]
//        public string NickName { get; set; }

//        public RunCollection()
//        {
//        }
//    }
//}
