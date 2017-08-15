using System;
using System.Diagnostics;
using System.ServiceModel;
using System.Windows.Media.Imaging;
using System.Threading.Tasks;
using Autodesk.Revit.UI;


namespace Lyrebird
{
    public class LBApp : IExternalApplication
    {
        static LBApp _thisApp = null;
        static UIControlledApplication _uiApp = null;
        static RibbonItem serverButton;

        ServiceHost serviceHost;
        Uri address;
        internal static UIApplication UIApp = null;

        bool serviceRunning = false;
        internal static LBHandler Handler = null;
        internal static ExternalEvent ExEvent = null;

        //public LBHandler LBHandle => Handler;
        //public ExternalEvent ExtEvent => ExEvent;

        public static LBApp Instance => _thisApp;

        public Result OnShutdown(UIControlledApplication application)
        {
            StopServer();
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            _thisApp = this;
            _uiApp = application;
            serviceRunning = false;
            address = new Uri(Properties.Settings.Default.BaseAddress + "/Revit" + _uiApp.ControlledApplication.VersionNumber);

            // create the toggle button
            PushButtonData pbd = new PushButtonData("Lyrebird Server", "Lyrebird\nServer", typeof(LBApp).Assembly.Location, "Lyrebird.ToggleCommand")
            {
                LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_Off.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()),
                ToolTip = "Lyrebird Server is current off.",
            };

            PushButtonData settingsPbd = new PushButtonData("Lyrebird Settings", "Lyrebird Settings", typeof(LBApp).Assembly.Location, "Lyrebird.SettingsCmd")
            {
                ToolTip = "Setup Lyrebird"
            };

            RibbonPanel panel = application.CreateRibbonPanel("Lyrebird");
            SplitButtonData sbd = new SplitButtonData("Lyrebird", "Lyrebird");
            SplitButton sb = panel.AddItem(sbd) as SplitButton;
            serverButton = sb.AddPushButton(pbd);
            sb.AddPushButton(settingsPbd);
            sb.IsSynchronizedWithCurrentItem = false;


            return Result.Succeeded;
        }

        public void Toggle(UIApplication uiApplication)
        {
            if (!serviceRunning)
            {
                if (StartServer())
                {
                    PushButton button = serverButton as PushButton;
                    BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    if (button != null)
                    {
                        button.LargeImage = bms;
                        button.ToolTip = "The Lyrebird Server is currently on and will accept requests for data or create objects.  Push button to toggle the server off.";
                    }
                    UIApp = uiApplication;
                    serviceRunning = true;
                }
                else
                {
                    UIApp = uiApplication;
                    serviceRunning = false;
                }
                
            }
            else
            {
                if (StopServer())
                {
                    PushButton button = serverButton as PushButton;
                    BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_Off.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    if (button != null)
                    {
                        button.LargeImage = bms;
                        button.ToolTip = "The Lyrebird Server is cu|rrently off and will not accept requests for data or create objects.  Push button to toggle the server on.";
                    }
                }
                UIApp = null;
                serviceRunning = false;
            }
        }

        private bool StartServer()
        {
            if (serviceHost == null || serviceHost.State != CommunicationState.Opened)
            {
                try
                {
                    serviceHost = new ServiceHost(typeof(LBService), address);
                    serviceHost.Description.Behaviors.Add(new System.ServiceModel.Discovery.ServiceDiscoveryBehavior());
                    NetNamedPipeBinding nnpb = new NetNamedPipeBinding();
                    nnpb.MaxReceivedMessageSize = int.MaxValue;
                    serviceHost.AddServiceEndpoint(new System.ServiceModel.Discovery.UdpDiscoveryEndpoint());
                    serviceHost.AddServiceEndpoint(typeof(ILyrebirdService), nnpb, string.Empty);
                    serviceHost.Open();
                    serviceRunning = true;
                    if (serviceRunning)
                    {
                        // Remove any tasks sent while the server was off
                        while (TaskContainer.Instance.HasTaskToPerform)
                        {
                            TaskContainer.Instance.DequeueTask();
                        }

                        // start the External Event
                        Handler = new LBHandler();
                        ExEvent = ExternalEvent.Create(Handler);
                    }
                    return true;
                }
                catch (AddressAlreadyInUseException)
                {
                    TaskDialog dlg = new TaskDialog("Lyrebird Error");
                    dlg.MainInstruction = "Error Opening Lyrebird Server";
                    dlg.MainContent = "A Lyrebird Server is already open in another instance of Revit.  Make sure you turn off all running Lyrebird Servers before toggling this one on.";
                    dlg.Show();
                    return false;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.Message);
                    serviceHost = null;
                    Handler = null;
                    ExEvent = null;
                    return false;
                }
            }

            return false;
        }

        private bool StopServer()
        {
            if (serviceHost != null)
            {
                try
                {
                    serviceHost.Abort();
                    serviceHost.Close();
                    serviceHost = null;
                    Handler = null;
                    ExEvent = null;
                    return true;
                }
                catch (Exception exception)
                {
                    Debug.WriteLine(exception.Message);
                    TaskDialog.Show("Error", exception.Message);
                    return false;
                }
            }
            else
                return false;
        }
    }
}
