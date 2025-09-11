using System.Windows;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class RssFeedsDialog : Window
    {
        public string FeedsText => txtFeeds.Text;
        public RssFeedsDialog(string initial)
        {
            InitializeComponent();
            txtFeeds.Text = initial ?? string.Empty;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
