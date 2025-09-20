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

        // OAuth controls
        private readonly ComboBox _cmbAuthMethod;
        private readonly ComboBox _cmbProvider;
        private readonly TextBox _txtClientId;
        private readonly PasswordBox _txtClientSecret;
        private readonly TextBox _txtTenant;
        private readonly TextBox _txtScope;
        private readonly TextBox _txtTokenEndpoint;
        private readonly TextBox _txtAccessToken;
        private readonly TextBox _txtRefreshToken;

        public EmailAccountsDialog(AppSettings settings)
        {
            Title = "Email Accounts";
            Width = 780; Height = 780; WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanMinimize;
            _settings = settings;

            var grid = new Grid { Margin = new Thickness(10) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(240) });
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
            for (int i = 0; i < 18; i++) form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            form.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int r = 0;
            form.Children.Add(new TextBlock { Text = "Display name:", Margin = new Thickness(0, 0, 6, 6) });
            _txtDisplay = new TextBox(); Grid.SetColumn(_txtDisplay, 1); form.Children.Add(_txtDisplay);

            r++; var lblUser = new TextBlock { Text = "Username:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblUser, r); form.Children.Add(lblUser);
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

            // OAuth section header
            r++; var hdrAuth = new TextBlock { Text = "Authentication", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 16, 0, 6) }; Grid.SetRow(hdrAuth, r); form.Children.Add(hdrAuth);

            r++; var lblMethod = new TextBlock { Text = "Auth Method:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblMethod, r); form.Children.Add(lblMethod);
            _cmbAuthMethod = new ComboBox { ItemsSource = new[] { "Password", "OAuth2" } }; Grid.SetRow(_cmbAuthMethod, r); Grid.SetColumn(_cmbAuthMethod, 1); form.Children.Add(_cmbAuthMethod);

            r++; var lblProvider = new TextBlock { Text = "OAuth Provider:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblProvider, r); form.Children.Add(lblProvider);
            _cmbProvider = new ComboBox { ItemsSource = new[] { "", "Google", "Microsoft", "Custom" } }; Grid.SetRow(_cmbProvider, r); Grid.SetColumn(_cmbProvider, 1); form.Children.Add(_cmbProvider);

            r++; var lblClientId = new TextBlock { Text = "Client Id:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblClientId, r); form.Children.Add(lblClientId);
            _txtClientId = new TextBox(); Grid.SetRow(_txtClientId, r); Grid.SetColumn(_txtClientId, 1); form.Children.Add(_txtClientId);

            r++; var lblClientSecret = new TextBlock { Text = "Client Secret:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblClientSecret, r); form.Children.Add(lblClientSecret);
            _txtClientSecret = new PasswordBox(); Grid.SetRow(_txtClientSecret, r); Grid.SetColumn(_txtClientSecret, 1); form.Children.Add(_txtClientSecret);

            r++; var lblTenant = new TextBlock { Text = "Tenant (MSFT):", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblTenant, r); form.Children.Add(lblTenant);
            _txtTenant = new TextBox(); Grid.SetRow(_txtTenant, r); Grid.SetColumn(_txtTenant, 1); form.Children.Add(_txtTenant);

            r++; var lblScope = new TextBlock { Text = "Scope(s):", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblScope, r); form.Children.Add(lblScope);
            _txtScope = new TextBox(); Grid.SetRow(_txtScope, r); Grid.SetColumn(_txtScope, 1); form.Children.Add(_txtScope);

            r++; var lblTokenEndpoint = new TextBlock { Text = "Token Endpoint:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblTokenEndpoint, r); form.Children.Add(lblTokenEndpoint);
            _txtTokenEndpoint = new TextBox(); Grid.SetRow(_txtTokenEndpoint, r); Grid.SetColumn(_txtTokenEndpoint, 1); form.Children.Add(_txtTokenEndpoint);

            r++; var lblAccessToken = new TextBlock { Text = "Access Token:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblAccessToken, r); form.Children.Add(lblAccessToken);
            _txtAccessToken = new TextBox(); Grid.SetRow(_txtAccessToken, r); Grid.SetColumn(_txtAccessToken, 1); form.Children.Add(_txtAccessToken);

            r++; var lblRefreshToken = new TextBlock { Text = "Refresh Token:", Margin = new Thickness(0, 8, 6, 6) }; Grid.SetRow(lblRefreshToken, r); form.Children.Add(lblRefreshToken);
            _txtRefreshToken = new TextBox(); Grid.SetRow(_txtRefreshToken, r); Grid.SetColumn(_txtRefreshToken, 1); form.Children.Add(_txtRefreshToken);

            // Bottom buttons
            var bottom = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
            var btnSave = new Button { Content = "Save", Width = 80, Height = 30, Padding = new Thickness(6, 2, 6, 2) };
            var btnClose = new Button { Content = "Close", Width = 80, Height = 30, Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(6, 2, 6, 2) };
            btnSave.Click += (s, e) => { SaveSelection(); _settings.Save(); _list.Items.Refresh(); };
            btnClose.Click += (s, e) => { DialogResult = true; Close(); };
            bottom.Children.Add(btnSave); bottom.Children.Add(btnClose);

            Grid.SetColumn(form, 2);
            grid.Children.Add(form);
            Grid.SetColumn(bottom, 2); Grid.SetRow(bottom, 21);
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

            _cmbAuthMethod.SelectedItem = string.IsNullOrEmpty(acc.AuthMethod) ? "Password" : acc.AuthMethod;
            _cmbProvider.SelectedItem = acc.OAuthProvider ?? "";
            _txtClientId.Text = acc.OAuthClientId ?? string.Empty;
            _txtClientSecret.Password = acc.OAuthClientSecret ?? string.Empty;
            _txtTenant.Text = acc.OAuthTenant ?? string.Empty;
            _txtScope.Text = acc.OAuthScope ?? string.Empty;
            _txtTokenEndpoint.Text = acc.OAuthTokenEndpoint ?? string.Empty;
            _txtAccessToken.Text = acc.OAuthAccessToken ?? string.Empty;
            _txtRefreshToken.Text = acc.OAuthRefreshToken ?? string.Empty;
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

            acc.AuthMethod = _cmbAuthMethod.SelectedItem as string ?? "Password";
            acc.OAuthProvider = _cmbProvider.SelectedItem as string ?? "";
            acc.OAuthClientId = _txtClientId.Text;
            acc.OAuthClientSecret = _txtClientSecret.Password;
            acc.OAuthTenant = _txtTenant.Text;
            acc.OAuthScope = _txtScope.Text;
            acc.OAuthTokenEndpoint = _txtTokenEndpoint.Text;
            acc.OAuthAccessToken = _txtAccessToken.Text;
            acc.OAuthRefreshToken = _txtRefreshToken.Text;
        }

        private void ClearDetails()
        {
            _txtDisplay.Text = _txtUser.Text = _txtImapHost.Text = _txtSmtpHost.Text = string.Empty;
            _txtPwd.Password = string.Empty;
            _txtImapPort.Text = ""; _txtSmtpPort.Text = "";
            _chkImapSsl.IsChecked = _chkSmtpSsl.IsChecked = true;

            _cmbAuthMethod.SelectedIndex = 0;
            _cmbProvider.SelectedIndex = 0;
            _txtClientId.Text = _txtTenant.Text = _txtScope.Text = _txtTokenEndpoint.Text = _txtAccessToken.Text = _txtRefreshToken.Text = string.Empty;
            _txtClientSecret.Password = string.Empty;
        }
    }
}
