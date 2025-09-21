using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Services
{
    public class EmailSyncService : IDisposable
    {
        public class MessageSummaryDto
        {
            public UniqueId Uid { get; set; }
            public string From { get; set; }
            public string Subject { get; set; }
            public DateTimeOffset Date { get; set; }
            public string DateDisplay => Date == default(DateTimeOffset) ? string.Empty : Date.LocalDateTime.ToString("g");
        }

        private class FolderCache
        {
            public DateTime LastUpdatedUtc { get; set; }
            public List<MessageSummaryDto> SummariesNewFirst { get; set; } = new List<MessageSummaryDto>();
            public Dictionary<UniqueId, string> HtmlByUid { get; set; } = new Dictionary<UniqueId, string>();
            public int LastKnownTotalCount { get; set; }
        }

        private readonly object _sync = new object();
        private readonly Dictionary<string, FolderCache> _cache = new Dictionary<string, FolderCache>();
        private AppSettings _settings;
        private Timer _timer;
        private bool _isRefreshing;

        private const int CacheLimit = 200;
        private const int PageSize = 25;

        public static EmailSyncService Current { get; } = new EmailSyncService();

        private EmailSyncService() { }

        public void Initialize(AppSettings settings)
        {
            _settings = settings ?? AppSettings.Load();
            // start timer: refresh every 5 minutes
            _timer = new Timer(async _ => await SafeRefreshAllAsync(), null, TimeSpan.Zero, TimeSpan.FromMinutes(5));
        }

        public async Task SafeRefreshAllAsync()
        {
            if (_isRefreshing) return;
            try
            {
                _isRefreshing = true;
                var accounts = (_settings?.EmailAccounts ?? new System.Collections.Generic.List<EmailAccount>()).ToList();
                foreach (var acc in accounts)
                {
                    try
                    {
                        await RefreshAccountFolderAsync(acc, "Inbox");
                        await RefreshAccountFolderAsync(acc, "Sent");
                        await RefreshAccountFolderAsync(acc, "Drafts");
                    }
                    catch { }
                }
            }
            finally { _isRefreshing = false; }
        }

        public async Task RefreshAccountFolderAsync(EmailAccount acc, string folder)
        {
            if (acc == null || string.IsNullOrWhiteSpace(folder)) return;
            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                client.CheckCertificateRevocation = false;
                await client.ConnectAsync(acc.ImapHost, acc.ImapPort, acc.ImapUseSsl);
                await AuthenticateImapAsync(client, acc);

                IMailFolder mailbox = client.Inbox;
                if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                await mailbox.OpenAsync(FolderAccess.ReadOnly);

                var total = mailbox.Count;
                int take = Math.Min(CacheLimit, total);
                int start = Math.Max(0, total - take);
                int end = Math.Max(0, total - 1);

                var summaries = total > 0
                    ? await mailbox.FetchAsync(start, end, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate)
                    : new List<IMessageSummary>();

                var list = summaries
                    .OrderByDescending(s => s.Date)
                    .Select(s => new MessageSummaryDto
                    {
                        Uid = s.UniqueId,
                        From = s.Envelope?.From?.FirstOrDefault()?.ToString() ?? string.Empty,
                        Subject = s.Envelope?.Subject ?? "(no subject)",
                        Date = s.Date
                    })
                    .ToList();

                var key = GetFolderKey(acc, folder);
                lock (_sync)
                {
                    if (!_cache.TryGetValue(key, out var fc))
                    {
                        fc = new FolderCache();
                        _cache[key] = fc;
                    }
                    fc.SummariesNewFirst = list;
                    fc.LastKnownTotalCount = total;
                    fc.LastUpdatedUtc = DateTime.UtcNow;
                }

                await client.DisconnectAsync(true);
            }
        }

        public List<MessageSummaryDto> GetSummaries(EmailAccount acc, string folder)
        {
            var key = GetFolderKey(acc, folder);
            lock (_sync)
            {
                if (_cache.TryGetValue(key, out var fc))
                    return fc.SummariesNewFirst.ToList();
            }
            return new List<MessageSummaryDto>();
        }

        public async Task EnsureOlderSummariesAsync(EmailAccount acc, string folder, int alreadyLoaded)
        {
            var key = GetFolderKey(acc, folder);
            FolderCache cache;
            lock (_sync) _cache.TryGetValue(key, out cache);

            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                client.CheckCertificateRevocation = false;
                await client.ConnectAsync(acc.ImapHost, acc.ImapPort, acc.ImapUseSsl);
                await AuthenticateImapAsync(client, acc);

                IMailFolder mailbox = client.Inbox;
                if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                await mailbox.OpenAsync(FolderAccess.ReadOnly);

                var total = mailbox.Count;
                int cached = cache?.SummariesNewFirst?.Count ?? 0;
                int alreadyCovered = Math.Min(CacheLimit, cached);
                int remainingOlder = Math.Max(0, total - alreadyCovered);
                if (remainingOlder <= 0)
                {
                    await client.DisconnectAsync(true);
                    return;
                }

                int fetchCount = Math.Min(PageSize, remainingOlder);
                int end = total - alreadyCovered - 1;
                int start = Math.Max(0, end - fetchCount + 1);

                var summaries = await mailbox.FetchAsync(start, end, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate);

                var older = summaries
                    .OrderByDescending(s => s.Date)
                    .Select(s => new MessageSummaryDto
                    {
                        Uid = s.UniqueId,
                        From = s.Envelope?.From?.FirstOrDefault()?.ToString() ?? string.Empty,
                        Subject = s.Envelope?.Subject ?? "(no subject)",
                        Date = s.Date
                    })
                    .ToList();

                lock (_sync)
                {
                    if (!_cache.TryGetValue(key, out cache))
                    {
                        cache = new FolderCache();
                        _cache[key] = cache;
                    }
                    cache.SummariesNewFirst.AddRange(older);
                    cache.LastKnownTotalCount = total;
                }

                await client.DisconnectAsync(true);
            }
        }

        public bool TryGetCachedBody(EmailAccount acc, string folder, UniqueId uid, out string html)
        {
            var key = GetFolderKey(acc, folder);
            lock (_sync)
            {
                if (_cache.TryGetValue(key, out var fc) && fc.HtmlByUid.TryGetValue(uid, out html))
                    return true;
            }
            html = null;
            return false;
        }

        public async Task<string> GetMessageBodyAsync(EmailAccount acc, string folder, UniqueId uid)
        {
            var key = GetFolderKey(acc, folder);
            if (TryGetCachedBody(acc, folder, uid, out var cached)) return cached;

            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                client.CheckCertificateRevocation = false;
                await client.ConnectAsync(acc.ImapHost, acc.ImapPort, acc.ImapUseSsl);
                await AuthenticateImapAsync(client, acc);

                IMailFolder mailbox = client.Inbox;
                if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                await mailbox.OpenAsync(FolderAccess.ReadOnly);

                var msg = await mailbox.GetMessageAsync(uid);
                var html = msg.HtmlBody ?? $"<pre>{System.Net.WebUtility.HtmlEncode(msg.TextBody ?? string.Empty)}</pre>";

                lock (_sync)
                {
                    if (!_cache.TryGetValue(key, out var fc))
                    {
                        fc = new FolderCache();
                        _cache[key] = fc;
                    }
                    fc.HtmlByUid[uid] = html;
                }

                await client.DisconnectAsync(true);
                return html;
            }
        }

        private string GetFolderKey(EmailAccount acc, string folder)
        {
            return ($"{acc?.Username}|{acc?.ImapHost}|{folder}").ToLowerInvariant();
        }

        // OAuth helper
        private static DateTime NowUtc => DateTime.UtcNow;
        private async Task AuthenticateImapAsync(MailKit.Net.Imap.ImapClient client, EmailAccount acc)
        {
            if (string.Equals(acc.AuthMethod, "OAuth2", StringComparison.OrdinalIgnoreCase))
            {
                var token = await EnsureAccessTokenAsync(acc);
                if (string.IsNullOrWhiteSpace(token))
                    throw new InvalidOperationException("No OAuth2 access token available.");
                var oauth2 = new MailKit.Security.SaslMechanismOAuth2(acc.Username, token);
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

            if (!string.IsNullOrWhiteSpace(acc.OAuthRefreshToken) && !string.IsNullOrWhiteSpace(acc.OAuthClientId))
            {
                var endpoint = ResolveTokenEndpoint(acc);
                if (!string.IsNullOrWhiteSpace(endpoint))
                {
                    var refreshed = await RefreshTokenAsync(endpoint, acc);
                    if (!string.IsNullOrWhiteSpace(refreshed))
                    {
                        acc.OAuthAccessToken = refreshed;
                        _settings?.Save();
                        return refreshed;
                    }
                }
            }

            // No interactive flow here (background). If token is expired and cannot be refreshed, return current value and let caller fail.
            return acc.OAuthAccessToken;
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
            catch { }
            return null;
        }

        public void Dispose()
        {
            try { _timer?.Dispose(); } catch { }
        }
    }
}
