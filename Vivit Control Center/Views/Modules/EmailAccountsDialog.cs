using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public class EmailAccountsDialog : Window
    {
        private readonly AppSettings _settings;
        private readonly ListBox _list;
        private readonly TextBox _txtDisplay;
        private readonly TextBox _txtUser;
        private readonly PasswordBox _txtPwd;
        private readonly TextBox _txtImapHost;
        private readonly TextBox _txtImapPort;
        private readonly CheckBox _chkImapSsl;
        private readonly TextBox _txtSmtpHost;
        private readonly TextBox _txtSmtpPort;
        private readonly CheckBox _chkSmtpSsl;

        public EmailAccountsDialog(AppSettings settings)
        {
            Title = "Email Accounts";
            Width = 680; Height = 420; WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanMinimize;
            _settings = settings;

            var grid = new Grid { Margin = new Thickness(10) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // List + buttons
            var left = new Grid();
            left.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            _list = new ListBox { DisplayMemberPath = "DisplayName" };
            _list.SelectionChanged += (s, e) => LoadSelection();
            _list.ItemsSource = _settings.EmailAccounts;
            Grid.SetRow(_list, 0);
            left.Children.Add(_list);
            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            var btnAdd = new Button { Content = "Add", Width = 90, Margin = new Thickness(0, 0, 6, 0) };
            var btnRemove = new Button { Content = "Remove", Width = 90 };
            btnAdd.Click += (s, e) => { var acc = new EmailAccount { DisplayName = "New Account" }; _settings.EmailAccounts.Add(acc); _list.Items.Refresh(); _list.SelectedItem = acc; };
            btnRemove.Click += (s, e) => { var sel = _list.SelectedItem as EmailAccount; if (sel != null) { _settings.EmailAccounts.Remove(sel); _list.Items.Refresh(); ClearDetails(); } };
            btnPanel.Children.Add(btnAdd); btnPanel.Children.Add(btnRemove);
            Grid.SetRow(btnPanel, 1);
            left.Children.Add(btnPanel);
            Grid.SetColumn(left, 0);
            grid.Children.Add(left);

            // Details form
            var form = new Grid();
            for (int i = 0; i < 9; i++) form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            form.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int r = 0;
            form.Children.Add(new TextBlock { Text = "Display name:", Margin = new Thickness(0, 0, 6, 6) });
            _txtDisplay = new TextBox(); Grid.SetColumn(_txtDisplay, 1); form.Children.Add(_txtDisplay);

            r++; Grid.SetRow(new TextBlock { Text = "Username:", Margin = new Thickness(0, 8, 6, 6) }, r);
            var lblUser = new TextBlock { Text = "Username:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblUser, r); form.Children.Add(lblUser);
            _txtUser = new TextBox(); Grid.SetRow(_txtUser, r); Grid.SetColumn(_txtUser, 1); form.Children.Add(_txtUser);

            r++; var lblPwd = new TextBlock { Text = "Password:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblPwd, r); form.Children.Add(lblPwd);
            _txtPwd = new PasswordBox(); Grid.SetRow(_txtPwd, r); Grid.SetColumn(_txtPwd, 1); form.Children.Add(_txtPwd);

            r++; var lblImapHost = new TextBlock { Text = "IMAP Host:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblImapHost, r); form.Children.Add(lblImapHost);
            _txtImapHost = new TextBox(); Grid.SetRow(_txtImapHost, r); Grid.SetColumn(_txtImapHost, 1); form.Children.Add(_txtImapHost);

            r++; var lblImapPort = new TextBlock { Text = "IMAP Port:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblImapPort, r); form.Children.Add(lblImapPort);
            _txtImapPort = new TextBox(); Grid.SetRow(_txtImapPort, r); Grid.SetColumn(_txtImapPort, 1); form.Children.Add(_txtImapPort);

            r++; var lblImapSsl = new TextBlock { Text = "IMAP SSL:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblImapSsl, r); form.Children.Add(lblImapSsl);
            _chkImapSsl = new CheckBox { VerticalAlignment = VerticalAlignment.Center }; Grid.SetRow(_chkImapSsl, r); Grid.SetColumn(_chkImapSsl, 1); form.Children.Add(_chkImapSsl);

            r++; var lblSmtpHost = new TextBlock { Text = "SMTP Host:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblSmtpHost, r); form.Children.Add(lblSmtpHost);
            _txtSmtpHost = new TextBox(); Grid.SetRow(_txtSmtpHost, r); Grid.SetColumn(_txtSmtpHost, 1); form.Children.Add(_txtSmtpHost);

            r++; var lblSmtpPort = new TextBlock { Text = "SMTP Port:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblSmtpPort, r); form.Children.Add(lblSmtpPort);
            _txtSmtpPort = new TextBox(); Grid.SetRow(_txtSmtpPort, r); Grid.SetColumn(_txtSmtpPort, 1); form.Children.Add(_txtSmtpPort);

            r++; var lblSmtpSsl = new TextBlock { Text = "SMTP SSL:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblSmtpSsl, r); form.Children.Add(lblSmtpSsl);
            _chkSmtpSsl = new CheckBox { VerticalAlignment = VerticalAlignment.Center }; Grid.SetRow(_chkSmtpSsl, r); Grid.SetColumn(_chkSmtpSsl, 1); form.Children.Add(_chkSmtpSsl);

            // Bottom buttons
            var bottom = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 150, 0, 0) };
            var btnSave = new Button { Content = "Save", Width = 80,Height=30, Padding = new Thickness(6,2,6,2) };
            var btnClose = new Button { Content = "Close", Width = 80, Height=30, Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(6,2,6,2) };
            btnSave.Click += (s, e) => { SaveSelection(); _settings.Save(); _list.Items.Refresh(); };
            btnClose.Click += (s, e) => { DialogResult = true; Close(); };
            bottom.Children.Add(btnSave); bottom.Children.Add(btnClose);

            Grid.SetColumn(form, 2);
            grid.Children.Add(form);
            Grid.SetColumn(bottom, 2); Grid.SetRow(bottom, 11);
            grid.Children.Add(bottom);

            Content = grid;
        }

        private void LoadSelection()
        {
            var acc = _list.SelectedItem as EmailAccount;
            if (acc == null) { ClearDetails(); return; }
            _txtDisplay.Text = acc.DisplayName;
            _txtUser.Text = acc.Username;
            _txtPwd.Password = acc.Password ?? string.Empty;
            _txtImapHost.Text = acc.ImapHost;
            _txtImapPort.Text = acc.ImapPort.ToString();
            _chkImapSsl.IsChecked = acc.ImapUseSsl;
            _txtSmtpHost.Text = acc.SmtpHost;
            _txtSmtpPort.Text = acc.SmtpPort.ToString();
            _chkSmtpSsl.IsChecked = acc.SmtpUseSsl;
        }

        private void SaveSelection()
        {
            var acc = _list.SelectedItem as EmailAccount;
            if (acc == null) return;
            acc.DisplayName = _txtDisplay.Text;
            acc.Username = _txtUser.Text;
            acc.Password = _txtPwd.Password;
            acc.ImapHost = _txtImapHost.Text;
            int.TryParse(_txtImapPort.Text, out var imapPort); acc.ImapPort = imapPort > 0 ? imapPort : 993;
            acc.ImapUseSsl = _chkImapSsl.IsChecked == true;
            acc.SmtpHost = _txtSmtpHost.Text;
            int.TryParse(_txtSmtpPort.Text, out var smtpPort); acc.SmtpPort = smtpPort > 0 ? smtpPort : 587;
            acc.SmtpUseSsl = _chkSmtpSsl.IsChecked == true;
        }

        private void ClearDetails()
        {
            _txtDisplay.Text = _txtUser.Text = _txtImapHost.Text = _txtSmtpHost.Text = string.Empty;
            _txtPwd.Password = string.Empty;
            _txtImapPort.Text = ""; _txtSmtpPort.Text = "";
            _chkImapSsl.IsChecked = _chkSmtpSsl.IsChecked = true;
        }
    }
}
