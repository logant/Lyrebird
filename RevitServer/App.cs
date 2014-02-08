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

namespace LMNA.Lyrebird
{
    public class RevitServerApp : IExternalApplication
    {
        bool serverActive;
        static RibbonItem serverButton;
        ServiceHost serviceHost;
        readonly Uri address = new Uri("net.pipe://localhost/LMNts/LyrebirdServer/Revit2014");
        bool disableButton;

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
            uicApp = application;
            application.Idling += OnIdling;
            
            _app = this;
            serverActive = Properties.Settings.Default.serverActive;
            
            try
            {
                BitmapSource bms;
                PushButtonData lyrebirdButton;
                StartServer();
                if (disableButton)
                {
                    bms = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.Lyrebird_On.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    lyrebirdButton = new PushButtonData("Lyrebird Server", "Lyrebird Server\nDisabled", typeof(RevitServerApp).Assembly.Location, "LMNA.Lyrebird.ServerToggle")
                    {
                        LargeImage = bms,
                        ToolTip = "The Lyrebird Server is currently disabled in this session of Revit.  This is most likely because you have more than one session of Revit and the server can only run in one.",
                    };
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
                }

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

                if (!found)
                {
                    // Create the panel
                    RibbonPanel utilitiesPanel = application.CreateRibbonPanel(tabName, "Utilities");
                    serverButton = utilitiesPanel.AddItem(lyrebirdButton);
                }
                else
                {
                    serverButton = panel.AddItem(lyrebirdButton);
                }

                if (disableButton)
                {
                    serverButton.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error2", ex.ToString());
            }

            return Result.Succeeded;
        }

        //TODO: Is this being used?  Do we need it?
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            Document doc = e.GetDocument();
            List<LyrebirdId> tempList = new List<LyrebirdId>();
            foreach(ElementId id in e.GetAddedElementIds())
            {
                Element elem = doc.GetElement(id);
                LyrebirdId lid = new LyrebirdId(elem.UniqueId, "Unknown");
                tempList.Add(lid);
            }
            createdIds = tempList;
        }

        private void OnIdling(object sender, IdlingEventArgs e)
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

        private void StartServer()
        {
            try
            {
                serviceHost = new ServiceHost(typeof(LyrebirdService), address);
                serviceHost.AddServiceEndpoint(typeof(ILyrebirdService), new NetNamedPipeBinding(), "LyrebirdService");
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

        private void ServiceOn()
        {
            // do something
        }
        private void ServiceOff()
        {
            // do something
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
                ServiceOff();
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
                ServiceOn();
            }
            Properties.Settings.Default.serverActive = serverActive;
            Properties.Settings.Default.Save();
        }
    }
}
