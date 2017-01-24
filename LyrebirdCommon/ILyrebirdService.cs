using System;
using System.Collections.Generic;
using System.ServiceModel;

namespace Lyrebird
{
    /// <summary>
    /// Lyrebird interface for Revit implementation
    /// </summary>
    [ServiceContract]
    public interface ILyrebirdService
    {
        [OperationContract]
        bool GetDocumentName(out string docName);

        [OperationContract]
        bool GetCategoryElements(ElementIdCategory eic, out List<string> categoryElements);

        [OperationContract]
        bool GetFamilies(out List<RevitObject> families);

        [OperationContract]
        bool GetTypeNames(string familyName, out List<string> typeNames);

        [OperationContract]
        bool GetParameters(string familyName, string typeName, out List<RevitParameter> parameters);

        [OperationContract]
        bool RecieveData(List<RevitObject> elements, Guid uniqueId, string nickName);
    }
}
