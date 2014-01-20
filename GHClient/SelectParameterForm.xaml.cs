using System;
using System.Collections.Generic;
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
using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird.GH
{
    /// <summary>
    /// Interaction logic for SelectParameterForm.xaml
    /// </summary>
    public partial class SelectParameterForm : Window
    {
        List<RevitParameter> parameters;

        LinearGradientBrush enterBrush;

        SetRevitDataForm parent;

        public SelectParameterForm(SetRevitDataForm _parent, List<RevitParameter> allParams, List<RevitParameter> selectedParams)
        {
            parameters = allParams;
            parent = _parent;
            InitializeComponent();

            // Add list items for each parameters
            listView.ItemsSource = parameters;
            listView.DisplayMemberPath = "ParameterName";

            // Preselected parameters
            if (selectedParams.Count > 0)
            {
                for (int i = 0; i < listView.Items.Count; i++)
                {
                    RevitParameter rp = listView.Items[i] as RevitParameter;
                    foreach (RevitParameter sp in selectedParams)
                    {
                        if (rp.ParameterName == sp.ParameterName)
                        {
                            listView.SelectedItems.Add(rp);
                        }
                    }
                }
            }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            List<RevitParameter> selectedParams = new List<RevitParameter>();
            try
            {
                foreach (RevitParameter item in listView.SelectedItems)
                {
                    try
                    {
                        RevitParameter rp = item;
                        selectedParams.Add(rp);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            parent.SelectedParameters = selectedParams;
            parent.AddControls();
            this.Close();
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
            okRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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
            cancelRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private LinearGradientBrush EnterBrush()
        {
            LinearGradientBrush brush = new LinearGradientBrush();
            brush.StartPoint = new System.Windows.Point(0, 0);
            brush.EndPoint = new System.Windows.Point(0, 1);
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 180, 180, 180), 0.0));
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 232, 232, 232), 1.0));
            return brush;
        }

        private void logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start(@"http://lmnts.lmnarchitects.com");
        }
    }
}
