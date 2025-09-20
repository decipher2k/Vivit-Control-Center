using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Vivit_Control_Center.Settings;

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
                    await client.ConnectAsync(_currentAccount.ImapHost, _currentAccount.ImapPort, _currentAccount.ImapUseSsl);
                    await client.AuthenticateAsync(_currentAccount.Username, _currentAccount.Password);
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
                    await client.ConnectAsync(_currentAccount.ImapHost, _currentAccount.ImapPort, _currentAccount.ImapUseSsl);
                    await client.AuthenticateAsync(_currentAccount.Username, _currentAccount.Password);
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
                    await client.ConnectAsync(_currentAccount.SmtpHost, _currentAccount.SmtpPort, _currentAccount.SmtpUseSsl ? SecureSocketOptions.StartTlsWhenAvailable : SecureSocketOptions.Auto);
                    await client.AuthenticateAsync(_currentAccount.Username, _currentAccount.Password);
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
    }
}
