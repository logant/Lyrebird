using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
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
                    //System.Windows.MessageBox.Show("I'm inside of Revit, in the LBService.LbAction method.");
                    Guid cmdGuid = Guid.Empty;
                    string path = string.Empty;
                    if (inputs.ContainsKey("CommandGuid"))
                        cmdGuid = (Guid)inputs["CommandGuid"];
                    if(inputs.ContainsKey("AssemblyPath"))
                        path = (string) inputs["AssemblyPath"];
                    //System.Windows.MessageBox.Show("Path: " + path + "\nGuid: " + cmdGuid.ToString());
                    if (cmdGuid != Guid.Empty && !string.IsNullOrEmpty(path))
                    {
                        //System.Windows.MessageBox.Show("I'm inside of Revit, in the LBService.LbAction method.\nPath: " + path + "\nGuid: " + cmdGuid.ToString());
                        var assembly = Assembly.LoadFrom(path);
                        Type[] types = null;
                        try
                        {
                            types = assembly.GetTypes();
                        }
                        catch (ReflectionTypeLoadException e)
                        {
                            types = e.Types.Where(t => t != null).ToArray();
                        }
                        //System.Windows.MessageBox.Show("Type array retrieved.\nNull?: " + (types == null).ToString());
                        
                        foreach (Type type in types)
                        {
                            //System.Windows.MessageBox.Show("In the types loop");
                            var prop =
                                type.GetProperty("CommandGuid", BindingFlags.Static | BindingFlags.Public);
                            if (prop != null)
                            {
                                //System.Windows.MessageBox.Show("found the property");
                                var val = prop.GetValue(null, null);
                                if (val == null)
                                    continue;
                                if ((Guid)val == cmdGuid)
                                {
                                    //System.Windows.MessageBox.Show("I'm about to try and launch the method.");
                                    
                                    // Load the RevitAPI libraries.
                                    string installPath = null;
                                    foreach (var reference in typeof(LBService).Assembly.GetReferencedAssemblies())
                                    {
                                        string refPath = Assembly.ReflectionOnlyLoad(reference.FullName).Location
                                            .ToLower();
                                        if (refPath.Contains("revitapi.dll") || refPath.Contains("revitapiui.dll"))
                                        {
                                            installPath = new FileInfo(refPath).DirectoryName;
                                            break;
                                        }
                                    }

                                    if (installPath == null)
                                        break;
                                    //System.Windows.MessageBox.Show("Testing:\n" + installPath);
                                    ConstructorInfo ctor = type.GetConstructor(new[] {typeof(string)});
                                    
                                    object inst = ctor.Invoke(new object[] {installPath});
                                    if (inst == null)
                                        break;
                                    //System.Windows.MessageBox.Show("Constructor instantiated");
                                    var method = type.GetMethod("Command");
                                    object[] parameters = {uiApp, inputs, new Dictionary<string, object>()};
                                    //System.Windows.MessageBox.Show("About to try and run the getdoc method");
                                    //System.Windows.MessageBox.Show(
//                                        "inst: " + (inst == null).ToString() +
//                                        "\nparameters: " + (inst == null).ToString() +
//                                        "\nmethod: " + (method == null).ToString());
                                    object result = method.Invoke(inst, parameters);
//                                    System.Windows.MessageBox.Show("We've run the get doc method!!");
                                    bool boolResult = (bool) result;
                                    if (boolResult)
                                        outputs = (Dictionary<string, object>) parameters[2];
                                    break;
                                }
                            }
                            else
                            {
                                System.Windows.MessageBox.Show("Type is not the one I want.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("I threw an exception.  :(\n" + ex.ToString());
                    //do nothing
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
                //outputs = new Dictionary<string, object> { { "docName", docName } };
            }
            
            return (outputs == null);
        }

        public bool GetApiPath(out List<string> apiDirectories)
        {
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    apiDirectories = new List<string>();
                    string installPath = null;
                    foreach (var reference in typeof(LBService).Assembly.GetReferencedAssemblies())
                    {
                        var location = Assembly.ReflectionOnlyLoad(reference.FullName).Location;
                        if (location != null)
                        {
                            string refPath = location.ToLower();
                            if (refPath.Contains("revitapi.dll") || refPath.Contains("revitapiui.dll"))
                                apiDirectories.Add(refPath);
                        }
                    }
                    apiDirectories = null;
                }
                catch
                {
                    apiDirectories = null;
                }
                Monitor.Wait(Locker, WAIT_TIMEOUT);
            }
            return (apiDirectories == null);
        }
    }
}                                                                              
