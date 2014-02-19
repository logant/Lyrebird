using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceModel;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB;
using LMNA.Lyrebird.LyrebirdCommon;
using Autodesk.Revit.ApplicationServices;

namespace LMNA.Lyrebird
{
    public class RevitServerApp : IExternalApplication
    {
        bool serverActive;
        static RibbonItem serverButton;
        ServiceHost serviceHost;
        readonly string addr = @"net.pipe://localhost/LMNts/LyrebirdServer/Revit";
        Uri address;
        bool disableButton;

        SettingsForm settingsForm = null;

        internal static UIApplication uiApp = null;
        UIControlledApplication uicApp;
        internal static RevitServerApp _app = null;

        static List<LyrebirdId> createdIds;

        public static List<LyrebirdId> CreatedIds
        {
            get { return createdIds; }
        }

        public static RevitServerApp Instance
        {
            get { return _app; }
        }

        public static UIApplication UIApp
        {
            get { return uiApp; }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= OnIdling;
            if (serviceHost != null)
            {
                try
                {
                    serviceHost.Close();
                }
                catch (Exception exception)
                {
                  Debug.WriteLine(exception.Message);
                }
            }
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            address = new Uri(addr + application.ControlledApplication.VersionNumber);
            uicApp = application;
            
            
            _app = this;
            serverActive = Properties.Settings.Default.defaultServerOn;

            // Create the button
            try
            {
                BitmapSource bms;
                PushButtonData lyrebirdButton;
                //StartServer();
                if (disableButton)
                {
                    bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    lyrebirdButton = new PushButtonData("Lyrebird Server", "Lyrebird Server\nDisabled", typeof(RevitServerApp).Assembly.Location, "LMNA.Lyrebird.ServerToggle")
                    {
                        LargeImage = bms,
                        ToolTip = "The Lyrebird Server is currently disabled in this session of Revit.  This is most likely because you have more than one session of Revit and the server can only run in one.",
                    };
                    Properties.Settings.Default.enableServer = false;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    if (serverActive)
                    {
                        bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                        lyrebirdButton = new PushButtonData("Lyrebird Server", "Lyrebird\nServer On", typeof(RevitServerApp).Assembly.Location, "LMNA.Lyrebird.ServerToggle")
                        {
                            LargeImage = bms,
                            ToolTip = "The Lyrebird Server currently on and will accept requests for data and can create objects.  Push button to toggle the server off.",
                        };
                        StartServer();
                        //ServiceOn();
                    }
                    else
                    {
                        bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_Off.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                        lyrebirdButton = new PushButtonData("Lyrebird Server", "Lyrebird\nServer Off", typeof(RevitServerApp).Assembly.Location, "LMNA.Lyrebird.ServerToggle")
                        {
                            LargeImage = bms,
                            ToolTip = "The Lyrebird Server is currently off and will not accept requests for data or create objects.  Push button to toggle the server on.",
                        };
                    }
                    Properties.Settings.Default.enableServer = true;
                    Properties.Settings.Default.Save();
                }

                // Settings button
                PushButtonData settingsButtonData = new PushButtonData("Lyrebird Settings", "Lyrebird Settings", typeof(RevitServerApp).Assembly.Location, "LMNA.Lyrebird.SettingsCmd")
                {
                    LargeImage = bms,
                    ToolTip = "Lyrebird Server settings.",
                };

                // Create the tab if necessary
                const string tabName = "LMN"; 
                Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
                Autodesk.Windows.RibbonTab tab = null;
                foreach (Autodesk.Windows.RibbonTab t in ribbon.Tabs)
                {
                    if (t.Id == tabName)
                    {
                        tab = t;
                    }
                }
                if (tab == null)
                {
                    application.CreateRibbonTab(tabName);
                }

                bool found = false;
                List<RibbonPanel> panels = application.GetRibbonPanels(tabName);
                RibbonPanel panel = null;
                foreach (RibbonPanel rp in panels)
                {
                    if (rp.Name == "Utilities")
                    {
                        panel = rp;
                        found = true;
                    }
                }

                SplitButtonData sbd = new SplitButtonData("Lyrebird", "Lyrebird");
                
                if (!found)
                {
                    // Create the panel
                    RibbonPanel utilitiesPanel = application.CreateRibbonPanel(tabName, "Utilities");

                    // Split button
                    SplitButton sb = utilitiesPanel.AddItem(sbd) as SplitButton;
                    serverButton = sb.AddPushButton(lyrebirdButton) as PushButton;
                    PushButton settingsButton = sb.AddPushButton(settingsButtonData) as PushButton;
                    sb.IsSynchronizedWithCurrentItem = false;
                }
                else
                {
                    SplitButton sb = panel.AddItem(sbd) as SplitButton;
                    serverButton = sb.AddPushButton(lyrebirdButton) as PushButton;
                    PushButton settingsButton = sb.AddPushButton(settingsButtonData) as PushButton;
                    sb.IsSynchronizedWithCurrentItem = false;
                }

                if (disableButton)
                {
                    serverButton.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
            }

            return Result.Succeeded;
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (serverActive)
            {
                if (uiApp == null)
                {
                    uiApp = sender as UIApplication;
                }
                e.SetRaiseWithoutDelay();

                if (!TaskContainer.Instance.HasTaskToPerform)
                {
                    return;
                }
                try
                {
                    var task = TaskContainer.Instance.DequeueTask();
                    task(uiApp);
                }
                catch (Exception ex)
                {
                    //TaskDialog.Show("Error at Idle", ex.Message);
                    if (uiApp != null) uiApp.Application.WriteJournalComment("Lyrebird Error: " + ex.ToString(), true);
                }
            }
        }

        public void ShowForm()
        {
            if (settingsForm == null || !settingsForm.IsVisible)
            {
                settingsForm = new SettingsForm(this);
                settingsForm.Show();
            }
        }

        private void StartServer()
        {
            if (serviceHost == null || serviceHost.State != CommunicationState.Opened)
            {
                try
                {
                    serviceHost = new ServiceHost(typeof(LyrebirdService), address);
                    NetNamedPipeBinding nnpb = new NetNamedPipeBinding();
                    nnpb.MaxReceivedMessageSize = int.MaxValue;
                    serviceHost.AddServiceEndpoint(typeof(ILyrebirdService), nnpb, "LyrebirdService");
                    serviceHost.Open();
                    if (serverActive)
                    {
                        ServiceOn();
                    }
                }
                catch (AddressAlreadyInUseException ex)
                {
                    Debug.WriteLine(ex.Message);
                    ServiceOff();
                    disableButton = true;
                }
                catch (AddressAccessDeniedException ex)
                {
                    // Couldn"t Open the Server
                    Debug.WriteLine(ex.Message);
                    ServiceOff();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.ToString());
                }
            }
        }

        private void StopServer()
        {
            if (serviceHost != null)
            {
                try
                {
                    serviceHost.Abort();
                    serviceHost.Close();
                    serviceHost = null;
                    ServiceOff();
                }
                catch (Exception exception)
                {
                    Debug.WriteLine(exception.Message);
                }
            }
        }

        private void ServiceOn()
        {
            // Remove any tasks queued while the server was off
            while (TaskContainer.Instance.HasTaskToPerform)
            {
                TaskContainer.Instance.DequeueTask();
            }

            // Start the idling event so it can receive from or send to Grasshopper
            uicApp.Idling += OnIdling;
        }
        private void ServiceOff()
        {
            // Stop the idling event so it can't receive from or send to Grasshopper
            uicApp.Idling -= OnIdling; 
        }

        public void Toggle()
        {
            if (serverActive)
            {
                serverActive = false;
                RibbonButton button = serverButton as RibbonButton;
                BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_Off.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                if (button != null)
                {
                    button.LargeImage = bms;
                    button.ItemText = "Lyrebird\nServer Off";
                    button.ToolTip = "The Lyrebird Server is currently off and will not accept requests for data or create objects.  Push button to toggle the server on.";
                    serverButton = button;
                }
                StopServer();
                //ServiceOff();
            }
            else
            {
                serverActive = true;
                RibbonButton rbutton = serverButton as RibbonButton;
                BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                if (rbutton != null)
                {
                    rbutton.LargeImage = bms;
                    rbutton.ItemText = "Lyrebird\nServer On";
                    rbutton.ToolTip = "The Lyrebird Server currently on and will accept requests for data and can create objects.  Push button to toggle the server off.";
                    serverButton = rbutton;
                }
                StartServer();
                //ServiceOn();
            }
            //Properties.Settings.Default.serverActive = serverActive;
            //Properties.Settings.Default.Save();
        }

        public void Disable()
        {
            RibbonButton rbutton = serverButton as RibbonButton;
            if (rbutton.Enabled)
            {
                BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                rbutton.LargeImage = bms;
                rbutton.ItemText = "Lyrebird Server\nDisabled";
                rbutton.ToolTip = "The Lyrebird Server is currently disabled in this session of Revit.  This is most likely because you have more than one session of Revit and the server can only run in one.";
                rbutton.Enabled = false;
                serverButton = rbutton;
                StopServer();
            }
        }

        public void Enable()
        {
            RibbonButton rbutton = serverButton as RibbonButton;
            if (!rbutton.Enabled)
            {
                if (serverActive)
                {
                    BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    if (rbutton != null)
                    {
                        rbutton.LargeImage = bms;
                        rbutton.ItemText = "Lyrebird\nServer On";
                        rbutton.ToolTip = "The Lyrebird Server currently on and will accept requests for data and can create objects.  Push button to toggle the server off.";
                        rbutton.Enabled = true;
                        serverButton = rbutton;
                    }
                    StartServer();
                }
                else
                {
                    BitmapSource bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    if (rbutton != null)
                    {
                        rbutton.LargeImage = bms;
                        rbutton.ItemText = "Lyrebird\nServer Off";
                        rbutton.ToolTip = "The Lyrebird Server is currently off and will not accept requests for data or create objects.  Push button to toggle the server on.";
                        rbutton.Enabled = true;
                        serverButton = rbutton;
                    }
                    StopServer();
                }
            }
            Toggle();
            Toggle();
        }
    }
}
