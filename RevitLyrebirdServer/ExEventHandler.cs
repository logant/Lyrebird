using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Lyrebird
{
    public class LBHandler : IExternalEventHandler
    {
        private Dictionary<string, object> inputs;
        private Dictionary<string, object> outputs;

        public Dictionary<string, object> Inputs
        {
            get { return inputs; }
            set { inputs = value; }
        }

        public Dictionary<string, object> Outputs
        {
            get { return outputs; }
            set { outputs = value; }
        }

        public void Execute(UIApplication app)
        {
            outputs = null;
            try
            {
                if (!inputs.ContainsKey("CommandGuid"))
                    return;

                Guid cmdGuid = (Guid)inputs["CommandGuid"];
                if (cmdGuid == Guid.Empty)
                    return;
                

                string dirPath = Properties.Settings.Default.LBComponentPath;
                Assembly assembly;
                Type cmdType;
                if (!GetCommandAssembly(cmdGuid, Directory.GetFiles(dirPath, "*.gha"), out assembly, out cmdType))
                    return;

                // Load the RevitAPI libraries.
                string installPath = null;
                foreach (var reference in typeof(LBHandler).Assembly.GetReferencedAssemblies())
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
                    return;
                     
                ConstructorInfo ctor = cmdType.GetConstructor(new[] {typeof(string)});

                object inst = ctor.Invoke(new object[] {installPath});
                if (inst == null)
                    return;

                var method = cmdType.GetMethod("Command");
                object[] parameters = {app, inputs, new Dictionary<string, object>()};

                object result = method.Invoke(inst, parameters);

                bool boolResult = (bool) result;
                if (boolResult)
                    outputs = (Dictionary<string, object>) parameters[2];
            }
            catch (Exception ex)
            {
                var dlg = new TaskDialog("Lyrebird Error")
                {
                    TitleAutoPrefix = false,
                    MainInstruction = "Error"
                };
                dlg.MainInstruction = ex.Message;
                dlg.Show();
            }
        }

        public string GetName()
        {
            return "Lyrebird Server";
        }

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
                        if ((Guid)val == cmdGuid)
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
