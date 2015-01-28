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

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird
{
    /// <summary>
    /// Interaction logic for SelectRunsForm.xaml
    /// </summary>
    public partial class SelectRunsForm : Window
    {
        List<RunCollection> runCollections;
        List<Runs> runs;
        RunCollection selectedRC;
        Runs selectedRun;
        UIDocument uiDoc;

        List<string> runNames;

        LinearGradientBrush enterBrush = null;

        public SelectRunsForm(List<RunCollection> rcs, UIDocument _uiDoc)
        {
            runCollections = rcs;
            uiDoc = _uiDoc;
            InitializeComponent();
            runNames = new List<string>();
            
            for (int i = 0; i < runCollections.Count; i++)
            {
                string value = runCollections[i].NickName + " : " + runCollections[i].ComponentGuid.ToString();
                runNames.Add(value);
            }

            // Setup the data.
            if (runCollections != null && runCollections.Count > 0)
            {
                guidComboBox.ItemsSource = runNames;
                //guidComboBox.DisplayMemberPath = "ComponentGuid";
                guidComboBox.SelectedIndex = 0;
            }
        }

        private void guidComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string value = e.AddedItems[0] as string;
            int selection = 0;
            for (int i = 0; i < runNames.Count; i++)
            {
                if (value == runNames[i])
                    selection = i;
            }
            RunCollection rc = runCollections[selection];
            selectedRC = rc;
            runs = selectedRC.Runs;

            runComboBox.ItemsSource = runs;
            runComboBox.DisplayMemberPath = "RunName";
            runComboBox.SelectedIndex = 0;
        }

        private void runComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                Runs r = e.AddedItems[0] as Runs;
                selectedRun = r;
                exampleLabel.Content = r.FamilyType;
                qtyLabel.Content = r.ElementIds.Count.ToString();
            }
            catch { }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start(@"http://lmnts.lmnarchitects.com");
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

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
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
            okRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // Change the Revit selection based on the selected run elements
            List<ElementId> elemSet = new List<ElementId>();
            foreach (int id in selectedRun.ElementIds)
            {
                ElementId eid = new ElementId(id);
                elemSet.Add(eid);
            }

            uiDoc.Selection.SetElementIds(elemSet);
            Close();
        }

        private LinearGradientBrush EnterBrush()
        {
            LinearGradientBrush brush = new LinearGradientBrush
            {
                StartPoint = new System.Windows.Point(0, 0),
                EndPoint = new System.Windows.Point(0, 1)
            };

            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 180, 180, 180), 0.0));
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 232, 232, 232), 1.0));
            return brush;
        }
    }
}
