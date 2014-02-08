using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird.GH
{
    /// <summary>
    /// Interaction logic for ParameterControl.xaml
    /// </summary>
    public partial class ParameterControl
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
            isTypeLabel.Content = parameter.IsType ? "Type Parameter" : "Instance Parameter";
            storageTypeLabel.Content = parameter.StorageType;
        }
    }
}
