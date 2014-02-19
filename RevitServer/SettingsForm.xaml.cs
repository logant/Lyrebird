using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LMNA.Lyrebird
{
    /// <summary>
    /// Interaction logic for SettingsForm.xaml
    /// </summary>
    public partial class SettingsForm : Window
    {
        LinearGradientBrush enterBrush = null;
        bool enableServer = true;
        bool supressWarning = false;
        bool defaultOn = true;
        int timeOut = 200;
        int createTimeout = 200;

        RevitServerApp app;

        public SettingsForm(RevitServerApp _app)
        {
            app = _app;
            InitializeComponent();
            enabledCheckBox.IsChecked = Properties.Settings.Default.enableServer;
            supressProfileCheckBox.IsChecked = Properties.Settings.Default.suppressWarning;
            defaultOnCheckBox.IsChecked = Properties.Settings.Default.defaultServerOn;
            timeOut = Properties.Settings.Default.infoTimeout;
            createTimeout = Properties.Settings.Default.serverTimeout;
            timeoutTextBox.Text = timeOut.ToString();
            createModifyTimeoutTextBox.Text = createTimeout.ToString();
        }

        private void enabledCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            enableServer = true;
        }

        private void enabledCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            enableServer = false;
        }

        private void supressProfileCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            supressWarning = true;
        }

        private void supressProfileCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            supressWarning = false;
        }

        private void defaultOnCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            defaultOn = true;
        }

        private void defaultOnCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            defaultOn = false;
        }

        private void timeoutTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                timeOut = Convert.ToInt32(timeoutTextBox.Text);
            }
            catch { }
        }

        private void createModifyTimeoutTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                createTimeout = Convert.ToInt32(createModifyTimeoutTextBox.Text);
            }
            catch { }
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // Change the settings
            Properties.Settings.Default.enableServer = enableServer;
            Properties.Settings.Default.suppressWarning = supressWarning;
            Properties.Settings.Default.defaultServerOn = defaultOn;
            Properties.Settings.Default.infoTimeout = timeOut;
            Properties.Settings.Default.serverTimeout = createTimeout;
            Properties.Settings.Default.Save();
            
            // If necessary, disable or enable the server
            if (enableServer)
            {
                app.Enable();
            }
            else
            {
                app.Disable();
            }

            Close();
        }

        private void okButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (enterBrush == null)
            {
                enterBrush = EnterBrush();
            }
            okRect.Fill = enterBrush;
        }

        private void okButton_MouseLeave(object sender, MouseEventArgs e)
        {
            okRect.Fill = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0));
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void cancelButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (enterBrush == null)
            {
                enterBrush = EnterBrush();
            }
            cancelRect.Fill = enterBrush;
        }

        private void cancelButton_MouseLeave(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0));
        }

        private void logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"http://lmnts.lmnarchitects.com");
            }
            catch { }
        }

        private LinearGradientBrush EnterBrush()
        {
            LinearGradientBrush brush = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(0, 1)
            };

            brush.GradientStops.Add(new GradientStop(Color.FromArgb(255, 180, 180, 180), 0.0));
            brush.GradientStops.Add(new GradientStop(Color.FromArgb(255, 232, 232, 232), 1.0));
            return brush;
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        
    }
}
