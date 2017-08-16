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
        /// <summary>
        /// General purpose command for both retrieving data from and modifying the Revig file. 
        /// </summary>
        /// <param name="inputs">Dictionary<string, object> for storing input information from GH</param>
        /// <param name="outputs">Dictionary<string, object> for sending information back to GH</param>
        /// <returns></returns>
        [OperationContract]
        bool LbAction(Dictionary<string, object> inputs, out Dictionary<string, object> outputs);
        
        #region Probably Garbage, Removal Pending
        /*
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

        [OperationContract]
        bool GetApiPath(out List<string> apiDirectory);
        */
        #endregion
    }
}
