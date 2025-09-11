using System.Windows.Controls;

namespace Vivit_Control_Center.Views
{
    public partial class ModuleView : UserControl
    {
        public ModuleView(string moduleName)
        {
            InitializeComponent();
            TitleText.Text = moduleName;
        }
    }
}