using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird.GH
{
    /// <summary>
    /// Interaction logic for SelectParameterForm.xaml
    /// </summary>
    public partial class SelectParameterForm
    {
        readonly List<RevitParameter> parameters;

        LinearGradientBrush enterBrush;

        readonly SetRevitDataForm parent;

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
                foreach (object t in listView.Items)
                {
                    RevitParameter rp = t as RevitParameter;
                    foreach (RevitParameter sp in selectedParams)
                    {
                        if (rp != null && rp.ParameterName == sp.ParameterName)
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
                    catch (Exception exception)
                    {
                      Debug.WriteLine(exception.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            parent.SelectedParameters = selectedParams;
            parent.AddControls();
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

        private void logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Process.Start(@"http://lmnts.lmnarchitects.com");
        }
    }
}
