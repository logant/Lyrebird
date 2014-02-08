using System;
using System.Collections.Generic;
using System.ServiceModel;

namespace LMNA.Lyrebird.LyrebirdCommon
{
    [ServiceContract]
    public interface ILyrebirdService
    {
        //[OperationContract]
        //bool CreateObjects(List<RevitObject> objects);

        [OperationContract]
        string GetDocumentName();

        [OperationContract]
        List<RevitObject> GetFamilyNames();

        [OperationContract]
        List<RevitParameter> GetParameters(RevitObject familyName, string typeName);

        [OperationContract]
        List<string> GetTypeNames(RevitObject familyName);

        [OperationContract]
        bool CreateOrModify(List<RevitObject> objects, Guid uniqueId);

    }
}
