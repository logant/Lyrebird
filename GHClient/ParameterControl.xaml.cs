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
    /// Interaction logic for ParameterControl.xaml
    /// </summary>
    public partial class ParameterControl : UserControl
    {
        RevitParameter parameter;

        public RevitParameter Parameter
        {
            get { return parameter; }
            set { parameter = value; }
        }

        public ParameterControl(RevitParameter rp)
        {
            parameter = rp;
            InitializeComponent();
            paramNameLabel.Content = parameter.ParameterName;
            if (parameter.IsType == true)
            {
                isTypeLabel.Content = "Type Parameter";
            }
            else
            {
                isTypeLabel.Content = "Instance Parameter";
            }
            storageTypeLabel.Content = parameter.StorageType;
        }
    }
}
