using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Threading;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    [ServiceBehavior]
    public class LBService : ILyrebirdService
    {
        private string currentDocName = "NULL";
        private static readonly object Locker = new object();
        private const int WAIT_TIMEOUT = 1000;

        /*
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
        */

        public bool LbAction(Dictionary<string, object> inputs, out Dictionary<string, object> outputs)
        {
            bool allowMessages = false;
            outputs = null;
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    if (string.IsNullOrEmpty(Properties.Settings.Default.LBComponentPath) ||
                        !Directory.Exists(Properties.Settings.Default.LBComponentPath))
                    {
                        TaskDialog dlg = new TaskDialog("Error");
                        dlg.TitleAutoPrefix = false;
                        dlg.MainInstruction = "Lyrebird Components Folder Not Set";
                        dlg.MainContent =
                            "An attempt was made to use Lyrebird before the Lyrebird Components folder was set. Please use the Lyrebird settings to specify where the Lyrebird components are. Hint, these are typically in the Grasshopper Libraries folder at %APPDATA%\\Grasshopper\\Libraries\\Lyrebird";
                        dlg.Show();
                        return false;
                    }
                    

                    if (!inputs.ContainsKey("CommandGuid"))
                        return false;
                    if (allowMessages)
                        System.Windows.MessageBox.Show("CommandGuid key is in the dictionary");

                    Guid cmdGuid = (Guid) inputs["CommandGuid"];
                    if (cmdGuid == Guid.Empty)
                        return false;
                    if (allowMessages)
                        System.Windows.MessageBox.Show("CommandGuid has been found");

                    string dirPath = Properties.Settings.Default.LBComponentPath;
                    Assembly assembly;
                    Type cmdType;
                    if (!GetCommandAssembly(cmdGuid, Directory.GetFiles(dirPath, "*.gha"), out assembly, out cmdType))
                        return false;

                    if (allowMessages)
                        System.Windows.MessageBox.Show("Assemblies have been found");

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
                        return false;
                    if (allowMessages)
                        System.Windows.MessageBox.Show("installPath has been found");

                    ConstructorInfo ctor = cmdType.GetConstructor(new[] {typeof(string)});

                    object inst = ctor.Invoke(new object[] {installPath});
                    if (inst == null)
                        return false;

                    if (allowMessages)
                        System.Windows.MessageBox.Show("type has been instantiated.");

                    var method = cmdType.GetMethod("Command");
                    object[] parameters = {uiApp, inputs, new Dictionary<string, object>()};

                    object result = method.Invoke(inst, parameters);

                    bool boolResult = (bool) result;
                    if (boolResult)
                        outputs = (Dictionary<string, object>) parameters[2];

                    //System.Windows.MessageBox.Show("'outputs' Keys: " + outputs.Count.ToString());
                    return true;

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

        /*
        public bool GetApiPath(out List<string> apiDirectories)
        {
            lock (Locker)
            {
                UIApplication uiApp = LBApp.UIApp;
                try
                {
                    apiDirectories = new List<string>();
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
        */

        public bool GetCommandAssembly(Guid cmdGuid, string[] filePaths, out Assembly assembly, out Type cmdType)
        {
            assembly = null;
            cmdType = null;
            foreach (string ghaPath in filePaths)
            {
                var currentAssembly = Assembly.LoadFrom(ghaPath);
                Type[] currentTypes = null;
                try
                {
                    currentTypes = currentAssembly.GetTypes();
                }
                catch (ReflectionTypeLoadException e)
                {
                    currentTypes = e.Types.Where(t => t != null).ToArray();
                }

                foreach (Type type in currentTypes)
                {

                    var prop =
                        type.GetProperty("CommandGuid");
                    if (prop != null)
                    {
                        var val = prop.GetValue(null, null);
                        if (val == null)
                            continue;
                        if ((Guid) val == cmdGuid)
                        {
                            assembly = currentAssembly;
                            cmdType = type;
                            break;
                        }
                    }
                }
                if (assembly != null)
                    break;
            }

            if (assembly == null || cmdType == null)
                return false;

            return true;
        }
    }
}