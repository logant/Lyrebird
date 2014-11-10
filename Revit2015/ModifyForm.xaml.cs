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
using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird
{
    /// <summary>
    /// Interaction logic for ModifyForm.xaml
    /// </summary>
    public partial class ModifyForm : Window
    {
        LyrebirdService svc;
        int selectedRun;
        List<Runs> runs;
        LinearGradientBrush enterBrush = null;

        public ModifyForm(LyrebirdService _svc, List<Runs> _runs)
        {
            svc = _svc;
            runs = _runs;

            InitializeComponent();

            // Find the runs
            runComboBox.ItemsSource = runs;
            runComboBox.DisplayMemberPath = "RunName";
            runComboBox.SelectedIndex = 0;
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void modifyButton_Click(object sender, RoutedEventArgs e)
        {
            svc.ModifyBehavior = 0;
            svc.RunId = selectedRun;
            Close();
        }

        private void modifyButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (enterBrush == null)
            {
                enterBrush = EnterBrush();
            }
            modifyRect.Fill = enterBrush;
        }

        private void modifyButton_MouseLeave(object sender, MouseEventArgs e)
        {
            modifyRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private void newButton_Click(object sender, RoutedEventArgs e)
        {
            svc.ModifyBehavior = 1;
            svc.RunId = selectedRun;
            Close();
        }

        private void newButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (enterBrush == null)
            {
                enterBrush = EnterBrush();
            }
            newRect.Fill = enterBrush;
        }

        private void newButton_MouseLeave(object sender, MouseEventArgs e)
        {
            newRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            svc.ModifyBehavior = 2;
            svc.RunId = selectedRun;
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
            cancelRect.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));
        }

        private void logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start(@"http://lmnts.lmnarchitects.com");
        }

        private void runComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Runs r = e.AddedItems[0] as Runs;
            selectedRun = r.RunId;
            exampleLabel.Content = r.FamilyType;
        }

        private LinearGradientBrush EnterBrush()
        {
            LinearGradientBrush brush = new LinearGradientBrush
            {
                StartPoint = new System.Windows.Point(0, 0),
                EndPoint = new System.Windows.Point(0, 1)
            };

            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 232, 232, 232), 0.0));
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 250, 250, 250), 0.15));
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 250, 250, 250), 0.85));
            brush.GradientStops.Add(new GradientStop(System.Windows.Media.Color.FromArgb(255, 232, 232, 232), 1.0));
            return brush;
        }
    }
}
