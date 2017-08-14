using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WinForms = System.Windows.Forms;

namespace Lyrebird
{
    /// <summary>
    /// Interaction logic for LyrebirdSettings.xaml
    /// </summary>
    public partial class LyrebirdSettings : Window
    {
        // Brushes for the button fills
        LinearGradientBrush eBrush = new LinearGradientBrush(
            Color.FromArgb(255, 195, 195, 195),
            Color.FromArgb(255, 245, 245, 245),
            new Point(0, 0),
            new Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));

        public LyrebirdSettings()
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(Properties.Settings.Default.LBComponentPath) &&
                System.IO.Directory.Exists(Properties.Settings.Default.LBComponentPath))
                pathTextBox.Text = Properties.Settings.Default.LBComponentPath;
        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            using (WinForms.FolderBrowserDialog dlg = new WinForms.FolderBrowserDialog())
            {
                if (!string.IsNullOrEmpty(Properties.Settings.Default.LBComponentPath) &&
                    System.IO.Directory.Exists(Properties.Settings.Default.LBComponentPath))
                    dlg.SelectedPath = Properties.Settings.Default.LBComponentPath;
                else
                    dlg.RootFolder = Environment.SpecialFolder.ApplicationData;

                WinForms.DialogResult result = dlg.ShowDialog();
                if (result == WinForms.DialogResult.OK)
                {
                    pathTextBox.Text = dlg.SelectedPath;
                }
            }
        }

        private void BrowseButton_OnMouseEnter(object sender, MouseEventArgs e)
        {
            browseRect.Fill = eBrush;
        }

        private void BrowseButton_OnMouseLeave(object sender, MouseEventArgs e)
        {
            browseRect.Fill = lBrush;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CancelButton_OnMouseEnter(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = eBrush;
        }

        private void CancelButton_OnMouseLeave(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = lBrush;
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.LBComponentPath = pathTextBox.Text;
            Properties.Settings.Default.Save();

            Close();
        }

        private void OkButton_OnMouseEnter(object sender, MouseEventArgs e)
        {
            okRect.Fill = eBrush;
        }

        private void OkButton_OnMouseLeave(object sender, MouseEventArgs e)
        {
            okRect.Fill = lBrush;
        }

        private void UIElement_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DragMove();
            }
            catch
            {
                // Ignore
            }
        }
    }
}
