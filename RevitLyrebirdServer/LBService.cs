﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    public class LBService : ILyrebirdService
    {
        private string currentDocName = "NULL";
        private static readonly object Locker = new object();
        private const int WAIT_TIMEOUT = 1000;

        public bool GetDocumentName(out string docName)
        {
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    docName = uiApp.ActiveUIDocument.Document.Title;
                }
                catch
                {
                    docName = "NOT FOUND";
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
            }
            return !string.IsNullOrEmpty(docName);
        }

        public bool GetCategoryElements(ElementIdCategory eic, out List<string> categoryElements)
        {
            throw new NotImplementedException();
        }

        public bool GetFamilies(out List<RevitObject> families)
        {
            throw new NotImplementedException();
        }

        public bool GetTypeNames(string familyName, out List<string> typeNames)
        {
            throw new NotImplementedException();
        }

        public bool GetParameters(string familyName, string typeName, out List<RevitParameter> parameters)
        {
            throw new NotImplementedException();
        }

        public bool RecieveData(List<RevitObject> elements, Guid uniqueId, string nickName)
        {
            throw new NotImplementedException();
        }

        public bool LbAction(Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            outputs = null;
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    //docName = uiApp.ActiveUIDocument.Document.Title;
                }
                catch
                {
                    //docName = "NOT FOUND";
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
                //outputs = new Dictionary<string, object> { { "docName", docName } };
            }
            
            return (outputs == null);
        }
    }
}                                                                              
