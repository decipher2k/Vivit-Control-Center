using System.Windows;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SettingsModule : BaseSimpleModule
    {
        // Hook to open Email Accounts dialog from Settings (optional shortcut)
        private void OpenEmailAccounts_Click(object sender, RoutedEventArgs e)
        {
            var s = AppSettings.Load();
            var dlg = new EmailAccountsDialog(s) { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() == true) s.Save();
        }
    }
}
