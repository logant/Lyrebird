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

        LinearGradientBrush enterBrush = null;

        public SelectRunsForm(List<RunCollection> rcs, UIDocument _uiDoc)
        {
            runCollections = rcs;
            uiDoc = _uiDoc;
            InitializeComponent();

            // Setup the data.
            // Find the runs
            guidComboBox.ItemsSource = runCollections;
            guidComboBox.DisplayMemberPath = "ComponentGuid";
            guidComboBox.SelectedIndex = 0;
        }

        private void guidComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RunCollection r = e.AddedItems[0] as RunCollection;
            selectedRC = r;
            runs = r.Runs;
            
            runComboBox.ItemsSource = runs;
            runComboBox.DisplayMemberPath = "RunName";
            runComboBox.SelectedIndex = 0;
        }

        private void runComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Runs r = e.AddedItems[0] as Runs;
            selectedRun = r;
            exampleLabel.Content = r.FamilyType;
            qtyLabel.Content = r.ElementIds.Count.ToString();
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
            Autodesk.Revit.UI.Selection.SelElementSet elemSet = Autodesk.Revit.UI.Selection.SelElementSet.Create();
            foreach (int id in selectedRun.ElementIds)
            {
                ElementId eid = new ElementId(id);
                Element elem = uiDoc.Document.GetElement(eid);
                elemSet.Add(elem);
            }

            uiDoc.Selection.Elements = elemSet;
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
