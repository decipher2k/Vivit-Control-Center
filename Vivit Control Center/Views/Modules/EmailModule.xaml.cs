using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Vivit_Control_Center.Settings;
using Vivit_Control_Center.Views.Modules.OAuth;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class EmailModule : BaseSimpleModule
    {
        private class MessageItem
        {
            public UniqueId Uid { get; set; }
            public string From { get; set; }
            public string Subject { get; set; }
            public string Date { get; set; }
        }

        private ObservableCollection<MessageItem> _messages = new ObservableCollection<MessageItem>();
        private AppSettings _settings;
        private EmailAccount _currentAccount;
        private int _pageSize = 25;
        private int _loadedCount = 0;

        public EmailModule()
        {
            InitializeComponent();
            _settings = AppSettings.Load();
            cmbAccounts.ItemsSource = _settings.EmailAccounts;
            lvMessages.ItemsSource = _messages;
        }

        public override Task PreloadAsync() => base.PreloadAsync();

        private void cmbAccounts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _currentAccount = cmbAccounts.SelectedItem as EmailAccount;
            _messages.Clear();
            _loadedCount = 0;
            txtStatus.Text = string.Empty;
        }

        private async void btnRefresh_Click(object sender, RoutedEventArgs e) => await LoadMessagesAsync(reset: true);

        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            expCompose.IsExpanded = true;
            txtTo.Text = string.Empty;
            txtSubject.Text = string.Empty;
            txtBody.Text = string.Empty;
        }

        private void btnAccounts_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new EmailAccountsDialog(_settings) { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() == true)
            {
                _settings = AppSettings.Load();
                cmbAccounts.ItemsSource = _settings.EmailAccounts;
            }
        }

        private async void lstFolders_SelectionChanged(object sender, SelectionChangedEventArgs e) => await LoadMessagesAsync(reset: true);

        private async Task LoadMessagesAsync(bool reset)
        {
            if (_currentAccount == null) { txtStatus.Text = "No account"; return; }
            var folder = (lstFolders.SelectedItem as ListBoxItem)?.Content?.ToString() ?? "Inbox";
            try
            {
                using (var client = new ImapClient())
                {
                    // Fix for CRL/OCSP reachability issues
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(_currentAccount.ImapHost, _currentAccount.ImapPort, _currentAccount.ImapUseSsl);

                    await AuthenticateImapAsync(client, _currentAccount);

                    IMailFolder mailbox = client.Inbox;
                    if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                    else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                    await mailbox.OpenAsync(FolderAccess.ReadOnly);

                    if (reset)
                    {
                        _messages.Clear();
                        _loadedCount = 0;
                    }

                    var total = mailbox.Count;
                    var start = Math.Max(0, total - _loadedCount - _pageSize);
                    var end = Math.Max(0, total - 1 - _loadedCount);
                    if (end >= start && total > 0)
                    {
                        var summaries = await mailbox.FetchAsync(start, end, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate);
                        foreach (var s in summaries.Reverse())
                        {
                            var dt = s.Date; // DateTimeOffset
                            var dateStr = dt == default(DateTimeOffset) ? string.Empty : dt.LocalDateTime.ToString("g");
                            _messages.Add(new MessageItem
                            {
                                Uid = s.UniqueId,
                                From = s.Envelope?.From?.FirstOrDefault()?.ToString() ?? "",
                                Subject = s.Envelope?.Subject ?? "(no subject)",
                                Date = dateStr
                            });
                        }
                        _loadedCount += (end - start + 1);
                    }

                    await client.DisconnectAsync(true);
                }
                txtStatus.Text = $"Loaded {_loadedCount} messages";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Error: {ex.Message}";
            }
        }

        private async void lvMessages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = lvMessages.SelectedItem as MessageItem;
            if (item == null || _currentAccount == null) return;
            var folder = (lstFolders.SelectedItem as ListBoxItem)?.Content?.ToString() ?? "Inbox";
            try
            {
                using (var client = new ImapClient())
                {
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(_currentAccount.ImapHost, _currentAccount.ImapPort, _currentAccount.ImapUseSsl);

                    await AuthenticateImapAsync(client, _currentAccount);

                    IMailFolder mailbox = client.Inbox;
                    if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                    else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                    await mailbox.OpenAsync(FolderAccess.ReadOnly);

                    var msg = await mailbox.GetMessageAsync(item.Uid);

                    txtMailHeader.Text = $"From: {msg.From}\nSubject: {msg.Subject}\nDate: {msg.Date.LocalDateTime:g}";
                    var html = msg.HtmlBody ?? $"<pre>{System.Net.WebUtility.HtmlEncode(msg.TextBody ?? string.Empty)}</pre>";
                    wbBody.NavigateToString($"<html><head><meta charset='utf-8'/></head><body>{html}</body></html>");

                    await client.DisconnectAsync(true);
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Error: {ex.Message}";
            }
        }

        private async void btnLoadMore_Click(object sender, RoutedEventArgs e) => await LoadMessagesAsync(reset: false);

        private async void btnSend_Click(object sender, RoutedEventArgs e)
        {
            if (_currentAccount == null) { txtStatus.Text = "No account"; return; }
            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(_currentAccount.DisplayName ?? _currentAccount.Username, _currentAccount.Username));
                foreach (var addr in (txtTo.Text ?? string.Empty).Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                    message.To.Add(MailboxAddress.Parse(addr.Trim()));
                message.Subject = txtSubject.Text ?? string.Empty;
                var bodyBuilder = new BodyBuilder { TextBody = txtBody.Text ?? string.Empty };
                message.Body = bodyBuilder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(_currentAccount.SmtpHost, _currentAccount.SmtpPort, _currentAccount.SmtpUseSsl ? SecureSocketOptions.StartTlsWhenAvailable : SecureSocketOptions.Auto);

                    await AuthenticateSmtpAsync(client, _currentAccount);

                    await client.SendAsync(message);
                    await client.DisconnectAsync(true);
                }

                txtStatus.Text = "Sent.";
                expCompose.IsExpanded = false;
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Send error: {ex.Message}";
            }
        }

        private void btnCancelCompose_Click(object sender, RoutedEventArgs e) => expCompose.IsExpanded = false;

        // ===== OAuth2 helpers =====
        private static DateTime NowUtc => DateTime.UtcNow;

        private async Task AuthenticateImapAsync(ImapClient client, EmailAccount acc)
        {
            if (string.Equals(acc.AuthMethod, "OAuth2", StringComparison.OrdinalIgnoreCase))
            {
                var token = await EnsureAccessTokenAsync(acc);
                if (string.IsNullOrWhiteSpace(token)) 
                    throw new InvalidOperationException("No OAuth2 access token available.");
                var oauth2 = new SaslMechanismOAuth2(acc.Username, token);
                await client.AuthenticateAsync(oauth2);
            }
            else
            {
                await client.AuthenticateAsync(acc.Username, acc.Password);
            }
        }

        private async Task AuthenticateSmtpAsync(SmtpClient client, EmailAccount acc)
        {
            if (string.Equals(acc.AuthMethod, "OAuth2", StringComparison.OrdinalIgnoreCase))
            {
                var token = await EnsureAccessTokenAsync(acc);
                if (string.IsNullOrWhiteSpace(token))
                    throw new InvalidOperationException("No OAuth2 access token available.");
                var oauth2 = new SaslMechanismOAuth2(acc.Username, token);
                await client.AuthenticateAsync(oauth2);
            }
            else
            {
                await client.AuthenticateAsync(acc.Username, acc.Password);
            }
        }

        private async Task<string> EnsureAccessTokenAsync(EmailAccount acc)
        {
            var margin = TimeSpan.FromMinutes(2);
            if (!string.IsNullOrWhiteSpace(acc.OAuthAccessToken) && acc.OAuthTokenExpiryUtc > NowUtc + margin)
            {
                return acc.OAuthAccessToken;
            }

            // Try refresh if possible
            if (!string.IsNullOrWhiteSpace(acc.OAuthRefreshToken) && !string.IsNullOrWhiteSpace(acc.OAuthClientId))
            {
                var endpoint = ResolveTokenEndpoint(acc);
                if (!string.IsNullOrWhiteSpace(endpoint))
                {
                    var refreshed = await RefreshTokenAsync(endpoint, acc);
                    if (!string.IsNullOrWhiteSpace(refreshed))
                    {
                        acc.OAuthAccessToken = refreshed;
                        _settings.Save();
                        return refreshed;
                    }
                }
            }

            // Interactive acquisition if configured
            if (!string.IsNullOrWhiteSpace(acc.OAuthClientId))
            {
                try
                {
                    var result = await OAuthHelper.AcquireTokensInteractiveAsync(acc);
                    if (result != null && !string.IsNullOrWhiteSpace(result.AccessToken))
                    {
                        acc.OAuthAccessToken = result.AccessToken;
                        acc.OAuthRefreshToken = string.IsNullOrWhiteSpace(result.RefreshToken) ? acc.OAuthRefreshToken : result.RefreshToken;
                        acc.OAuthTokenExpiryUtc = result.ExpiryUtc;
                        _settings.Save();
                        return acc.OAuthAccessToken;
                    }
                }
                catch (Exception ex)
                {
                    txtStatus.Text = $"OAuth error: {ex.Message}";
                }
            }

            return acc.OAuthAccessToken; // may be null/empty
        }

        private static string ResolveTokenEndpoint(EmailAccount acc)
        {
            if (!string.IsNullOrWhiteSpace(acc.OAuthTokenEndpoint)) return acc.OAuthTokenEndpoint;
            var provider = (acc.OAuthProvider ?? "").Trim().ToLowerInvariant();
            if (provider == "google") return "https://oauth2.googleapis.com/token";
            if (provider == "microsoft")
            {
                var tenant = string.IsNullOrWhiteSpace(acc.OAuthTenant) ? "common" : acc.OAuthTenant.Trim();
                return $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token";
            }
            return null;
        }

        private static async Task<string> RefreshTokenAsync(string tokenEndpoint, EmailAccount acc)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    var content = new StringContent(
                        $"client_id={Uri.EscapeDataString(acc.OAuthClientId ?? string.Empty)}&" +
                        (!string.IsNullOrEmpty(acc.OAuthClientSecret) ? $"client_secret={Uri.EscapeDataString(acc.OAuthClientSecret)}&" : string.Empty) +
                        $"grant_type=refresh_token&refresh_token={Uri.EscapeDataString(acc.OAuthRefreshToken ?? string.Empty)}&" +
                        (!string.IsNullOrWhiteSpace(acc.OAuthScope) ? $"scope={Uri.EscapeDataString(acc.OAuthScope)}&" : string.Empty),
                        Encoding.UTF8, "application/x-www-form-urlencoded");
                    var resp = await http.PostAsync(tokenEndpoint, content);
                    var body = await resp.Content.ReadAsStringAsync();
                    if (!resp.IsSuccessStatusCode)
                        throw new Exception($"Token refresh failed: {resp.StatusCode} {body}");

                    // Parse JSON lightweight
                    var access = Regex.Match(body, "\\\"access_token\\\"\\s*:\\s*\\\"(.*?)\\\"").Groups[1].Value;
                    var expiresStr = Regex.Match(body, "\\\"expires_in\\\"\\s*:\\s*(\\d+)").Groups[1].Value;
                    var refresh = Regex.Match(body, "\\\"refresh_token\\\"\\s*:\\s*\\\"(.*?)\\\"").Groups[1].Value;
                    if (!string.IsNullOrEmpty(access))
                    {
                        int expiresSec = 3600;
                        int.TryParse(expiresStr, out expiresSec);
                        acc.OAuthTokenExpiryUtc = NowUtc.AddSeconds(expiresSec - 60);
                        if (!string.IsNullOrEmpty(refresh)) acc.OAuthRefreshToken = refresh;
                        return access;
                    }
                }
            }
            catch
            {
            }
            return null;
        }
    }
}
