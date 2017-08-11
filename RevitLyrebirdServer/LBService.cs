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
                    Guid cmdGuid = Guid.Empty;
                    string path = string.Empty;
                    if (inputs.ContainsKey("CommandGuid"))
                        cmdGuid = (Guid)inputs["CommandGuid"];
                    if(inputs.ContainsKey("AssemblyPath"))
                        path = (string) inputs["AssemblyPath"];

                    if (cmdGuid != Guid.Empty && string.IsNullOrEmpty(path))
                    {
                        var assembly = Assembly.LoadFrom(path);
                        foreach (Type type in assembly.GetTypes())
                        {
                            var prop =
                                type.GetProperty("CommandGuid", BindingFlags.Static | BindingFlags.Public);
                            if (prop != null)
                            {
                                var val = prop.GetValue(null, null);
                                if (val == null)
                                    continue;
                                if ((Guid)val == cmdGuid)
                                {
                                    var method = type.GetMethod("Command", BindingFlags.Static);
                                    object[] parameters = {uiApp, inputs, null};
                                    object result = method.Invoke(null, parameters);
                                    bool boolResult = (bool) result;
                                    if (boolResult)
                                        outputs = (Dictionary<string, object>) parameters[2];
                                    break;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    //do nothing
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
                //outputs = new Dictionary<string, object> { { "docName", docName } };
            }
            
            return (outputs == null);
        }
    }
}                                                                              
