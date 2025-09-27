using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
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
            public bool IsUnread { get; set; }
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

        private static string CacheRoot
        {
            get
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VivitControlCenter", "mailcache");
                try { Directory.CreateDirectory(dir); } catch { }
                return dir;
            }
        }
        private static string GetAccountFolder(EmailAccount acc)
        {
            var key = ($"{acc?.Username}|{acc?.ImapHost}").ToLowerInvariant();
            var safe = Convert.ToBase64String(Encoding.UTF8.GetBytes(key));
            return Path.Combine(CacheRoot, safe);
        }
        private static string GetFolderDir(EmailAccount acc, string folder)
        {
            var dir = Path.Combine(GetAccountFolder(acc), (folder ?? "Inbox").ToLowerInvariant());
            try { Directory.CreateDirectory(dir); } catch { }
            return dir;
        }
        private static string GetSummariesPath(EmailAccount acc, string folder) => Path.Combine(GetFolderDir(acc, folder), "summaries.txt");
        private static string GetBodiesDir(EmailAccount acc, string folder)
        {
            var dir = Path.Combine(GetFolderDir(acc, folder), "bodies");
            try { Directory.CreateDirectory(dir); } catch { }
            return dir;
        }
        private static string GetBodyFile(EmailAccount acc, string folder, UniqueId uid)
        {
            // Use UID numeric Id component for filename safety
            var name = uid.Id.ToString() + ".html";
            return Path.Combine(GetBodiesDir(acc, folder), name);
        }

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
                    ? await mailbox.FetchAsync(start, end, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate | MessageSummaryItems.Flags)
                    : new List<IMessageSummary>();

                var list = summaries
                    .OrderByDescending(s => s.Date)
                    .Select(s => new MessageSummaryDto
                    {
                        Uid = s.UniqueId,
                        From = s.Envelope?.From?.FirstOrDefault()?.ToString() ?? string.Empty,
                        Subject = s.Envelope?.Subject ?? "(no subject)",
                        Date = s.Date,
                        IsUnread = s.Flags.HasValue ? !s.Flags.Value.HasFlag(MessageFlags.Seen) : true // assume unread if unknown
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

                // Persist cache to disk
                try { SaveSummaries(acc, folder, list); } catch { }

                await client.DisconnectAsync(true);
            }
        }

        public async Task TryMarkAsSeenAsync(EmailAccount acc, string folder, UniqueId uid)
        {
            try
            {
                using (var client = new ImapClient())
                {
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(acc.ImapHost, acc.ImapPort, acc.ImapUseSsl);
                    await AuthenticateImapAsync(client, acc);

                    IMailFolder mailbox = client.Inbox;
                    if (string.Equals(folder, "Sent", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Sent);
                    else if (string.Equals(folder, "Drafts", StringComparison.OrdinalIgnoreCase)) mailbox = client.GetFolder(SpecialFolder.Drafts);
                    await mailbox.OpenAsync(FolderAccess.ReadWrite);

                    await mailbox.AddFlagsAsync(uid, MessageFlags.Seen, true);

                    // update cache entry if present
                    var key = GetFolderKey(acc, folder);
                    lock (_sync)
                    {
                        if (_cache.TryGetValue(key, out var fc))
                        {
                            var m = fc.SummariesNewFirst.FirstOrDefault(x => x.Uid == uid);
                            if (m != null) m.IsUnread = false;
                        }
                    }

                    await client.DisconnectAsync(true);
                }
            }
            catch { }
        }

        public List<MessageSummaryDto> GetSummaries(EmailAccount acc, string folder)
        {
            var key = GetFolderKey(acc, folder);
            lock (_sync)
            {
                if (_cache.TryGetValue(key, out var fc))
                    return fc.SummariesNewFirst.ToList();
            }
            // Try load from disk on-demand
            try
            {
                var disk = LoadSummaries(acc, folder);
                if (disk != null && disk.Count > 0)
                {
                    lock (_sync)
                    {
                        var fc = new FolderCache { SummariesNewFirst = disk, LastKnownTotalCount = disk.Count, LastUpdatedUtc = DateTime.UtcNow };
                        _cache[key] = fc;
                    }
                    return disk.ToList();
                }
            }
            catch { }
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

                var summaries = await mailbox.FetchAsync(start, end, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate | MessageSummaryItems.Flags);

                var older = summaries
                    .OrderByDescending(s => s.Date)
                    .Select(s => new MessageSummaryDto
                    {
                        Uid = s.UniqueId,
                        From = s.Envelope?.From?.FirstOrDefault()?.ToString() ?? string.Empty,
                        Subject = s.Envelope?.Subject ?? "(no subject)",
                        Date = s.Date,
                        IsUnread = s.Flags.HasValue ? !s.Flags.Value.HasFlag(MessageFlags.Seen) : false
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

                // Update disk cache with appended items
                try
                {
                    var all = GetSummaries(acc, folder);
                    SaveSummaries(acc, folder, all);
                }
                catch { }

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
            // Try disk
            try
            {
                var path = GetBodyFile(acc, folder, uid);
                if (File.Exists(path))
                {
                    html = File.ReadAllText(path, Encoding.UTF8);
                    // Populate memory cache for faster next access
                    lock (_sync)
                    {
                        if (!_cache.TryGetValue(key, out var fc)) { fc = new FolderCache(); _cache[key] = fc; }
                        fc.HtmlByUid[uid] = html;
                    }
                    return true;
                }
            }
            catch { }
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

                // Persist body to disk
                try
                {
                    var path = GetBodyFile(acc, folder, uid);
                    File.WriteAllText(path, html, Encoding.UTF8);
                }
                catch { }

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

        // ===== Disk cache helpers =====
        private static void SaveSummaries(EmailAccount acc, string folder, List<MessageSummaryDto> list)
        {
            var path = GetSummariesPath(acc, folder);
            try
            {
                var sb = new StringBuilder();
                // header with timestamp
                sb.AppendLine(DateTime.UtcNow.Ticks.ToString());
                foreach (var m in list)
                {
                    // uidId|ticks|isUnread|fromBase64|subjBase64
                    var uidId = m.Uid.Id.ToString();
                    var ticks = m.Date.UtcTicks.ToString();
                    var unread = m.IsUnread ? "1" : "0";
                    var from64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(m.From ?? string.Empty));
                    var subj64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(m.Subject ?? string.Empty));
                    sb.AppendLine(string.Join("|", new[] { uidId, ticks, unread, from64, subj64 }));
                }
                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        private static List<MessageSummaryDto> LoadSummaries(EmailAccount acc, string folder)
        {
            var path = GetSummariesPath(acc, folder);
            if (!File.Exists(path)) return new List<MessageSummaryDto>();
            var list = new List<MessageSummaryDto>();
            try
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                for (int i = 1; i < lines.Length; i++) // skip header
                {
                    var line = lines[i];
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('|');
                    if (parts.Length < 5) continue;
                    uint uidId; long ticks; bool unread;
                    if (!uint.TryParse(parts[0], out uidId)) continue;
                    if (!long.TryParse(parts[1], out ticks)) ticks = 0L;
                    unread = parts[2] == "1";
                    string from = SafeB64(parts[3]);
                    string subj = SafeB64(parts[4]);
                    try
                    {
                        // UniqueId ctor that accepts uint id
                        var uid = new UniqueId(uidId);
                        var dto = new MessageSummaryDto
                        {
                            Uid = uid,
                            From = from,
                            Subject = subj,
                            Date = ticks > 0 ? new DateTimeOffset(new DateTime(ticks, DateTimeKind.Utc)) : default(DateTimeOffset),
                            IsUnread = unread
                        };
                        list.Add(dto);
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        private static string SafeB64(string s)
        {
            try { return Encoding.UTF8.GetString(Convert.FromBase64String(s ?? string.Empty)); } catch { return string.Empty; }
        }
    }
}
