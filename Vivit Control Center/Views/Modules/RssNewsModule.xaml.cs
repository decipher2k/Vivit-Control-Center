using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml.Linq;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class RssNewsModule : BaseSimpleModule
    {
        private readonly ObservableCollection<RssArticle> _articles = new ObservableCollection<RssArticle>();
        private bool _loaded;
        private AppSettings _settings;
        private static readonly HttpClient _http = new HttpClient();

        private DispatcherTimer _refreshTimer;
        private bool _refreshInProgress;

        // central image http client
        private static readonly HttpClient _imageHttp;
        private static readonly SemaphoreSlim _thumbConcurrency = new SemaphoreSlim(4);

        static RssNewsModule()
        {
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | (SecurityProtocolType)3072; } catch { }
            _imageHttp = new HttpClient(new HttpClientHandler
            {
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                MaxAutomaticRedirections = 5,
                UseCookies = false
            });
            _imageHttp.Timeout = TimeSpan.FromSeconds(15);
            _imageHttp.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36");
            _imageHttp.DefaultRequestHeaders.Accept.ParseAdd("image/jpeg,image/png,image/gif,image/bmp,image/*;q=0.8,*/*;q=0.5");
        }

        public RssNewsModule()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += async (_, __) => { await EnsureLoadedAsync(); await InitArticleWebViewAsync(); };
            Unloaded += (_, __) => { try { _refreshTimer?.Stop(); } catch { } };
        }

        private Microsoft.Web.WebView2.Wpf.WebView2 GetArticleWebView() => this.FindName("wbArticle") as Microsoft.Web.WebView2.Wpf.WebView2;

        private async Task InitArticleWebViewAsync()
        {
            try
            {
                var wv = GetArticleWebView();
                if (wv == null) return;
                await wv.EnsureCoreWebView2Async();
                var s = wv.CoreWebView2.Settings;
                s.IsScriptEnabled = false;
                s.AreDefaultScriptDialogsEnabled = false;
                s.AreDefaultContextMenusEnabled = true;
                s.IsWebMessageEnabled = false;
                s.IsGeneralAutofillEnabled = false;
                s.IsPasswordAutosaveEnabled = false;
            }
            catch { }
        }

        private async Task EnsureLoadedAsync()
        {
            if (_loaded) return;
            _loaded = true;
            _settings = AppSettings.Load();
            await LoadFeedsAsync();
            EnsureRefreshTimer();
            try { _ = Dispatcher.InvokeAsync(async () => { await Task.Delay(2000).ConfigureAwait(false); await RefreshNowAsync().ConfigureAwait(false); }); } catch { }
        }

        public override async Task PreloadAsync()
        {
            await EnsureLoadedAsync();
            await base.PreloadAsync();
        }

        private void EnsureRefreshTimer()
        {
            if (_refreshTimer == null)
            {
                _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(5) };
                _refreshTimer.Tick += async (s, e) => await RefreshNowAsync();
            }
            try { _refreshTimer.Start(); } catch { }
        }

        private async Task RefreshNowAsync()
        {
            if (_refreshInProgress) return;
            _refreshInProgress = true;
            try { await LoadFeedsAsync(); }
            catch { }
            finally { _refreshInProgress = false; }
        }

        public override void SetVisible(bool visible)
        {
            base.SetVisible(visible);
            if (visible)
            {
                try { _refreshTimer?.Start(); } catch { }
                _ = RefreshNowAsync();
            }
            else
            {
                try { _refreshTimer?.Start(); } catch { }
            }
        }

        private async Task LoadFeedsAsync()
        {
            var feeds = (_settings.RssFeeds ?? new System.Collections.Generic.List<string>())
                .Where(f => !string.IsNullOrWhiteSpace(f))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (feeds.Count == 0)
            {
                lstArticles.ItemsSource = null;
                return;
            }

            int totalTarget = _settings.RssMaxArticles > 0 ? _settings.RssMaxArticles : 60;

            // Fetch feeds in parallel
            var tasks = feeds.Select(f => FetchFeedArticlesAsync(f, totalTarget)).ToArray();
            System.Collections.Generic.List<RssArticle>[] perFeedLists;
            try { perFeedLists = await Task.WhenAll(tasks).ConfigureAwait(false); }
            catch { perFeedLists = tasks.Where(t => t.IsCompleted).Select(t => t.Result).ToArray(); }

            // Round-robin merge up to target
            var final = new System.Collections.Generic.List<RssArticle>(totalTarget);
            bool progress = true;
            var lists = perFeedLists.ToList();
            while (final.Count < totalTarget && progress)
            {
                progress = false;
                for (int i = 0; i < lists.Count && final.Count < totalTarget; i++)
                {
                    var list = lists[i];
                    if (list == null || list.Count == 0) continue;
                    final.Add(list[0]);
                    list.RemoveAt(0);
                    progress = true;
                }
            }

            // Update UI on dispatcher
            await Dispatcher.InvokeAsync(() =>
            {
                var view = new ListCollectionView(final);
                view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(RssArticle.Source)));
                lstArticles.ItemsSource = view;
            });

            // kick off thumbnail loading for visible items
            foreach (var art in final)
            {
                QueueThumbnailForArticle(art);
            }
        }

        private async Task<System.Collections.Generic.List<RssArticle>> FetchFeedArticlesAsync(string feed, int softLimit)
        {
            var result = new System.Collections.Generic.List<RssArticle>();
            try
            {
                var bytes = await _http.GetByteArrayAsync(feed).ConfigureAwait(false);
                var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(bytes));
                if (xml.Root == null) return result;

                if (xml.Root.Name.LocalName.Equals("rss", StringComparison.OrdinalIgnoreCase) ||
                    xml.Root.Elements().Any(e => e.Name.LocalName == "channel"))
                {
                    var channel = xml.Root.Element("channel") ?? xml.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "channel");
                    var sourceTitle = channel?.Element("title")?.Value?.Trim() ?? feed;
                    foreach (var item in channel?.Elements().Where(e => e.Name.LocalName == "item") ?? Enumerable.Empty<XElement>())
                    {
                        var art = BuildArticleFromRssItem(item, sourceTitle, feed);
                        if (art != null) result.Add(art);
                        if (result.Count >= softLimit) break;
                    }
                }
                else if (xml.Root.Name.LocalName.Equals("feed", StringComparison.OrdinalIgnoreCase))
                {
                    var sourceTitle = xml.Root.Element(xml.Root.GetDefaultNamespace() + "title")?.Value?.Trim() ?? feed;
                    foreach (var entry in xml.Root.Elements().Where(e => e.Name.LocalName == "entry"))
                    {
                        var art = BuildArticleFromAtomEntry(entry, sourceTitle, feed);
                        if (art != null) result.Add(art);
                        if (result.Count >= softLimit) break;
                    }
                }
            }
            catch { }

            return result.OrderByDescending(a => a.PublishDate).ToList();
        }

        private RssArticle BuildArticleFromRssItem(XElement item, string sourceTitle, string feedUrl)
        {
            try
            {
                string title = item.Element("title")?.Value?.Trim();
                string link = item.Element("link")?.Value?.Trim();
                string desc = item.Element("description")?.Value ??
                              item.Elements().FirstOrDefault(e => e.Name.LocalName == "encoded")?.Value ??
                              item.Element("content")?.Value ?? string.Empty;
                DateTime pub = ParseDate(item.Element("pubDate")?.Value) ?? DateTime.UtcNow;
                string enclosure = ResolveUrl(feedUrl, ExtractEnclosureUrl(item));
                string imgUrl = ResolveUrl(feedUrl, ExtractImageUrl(item) ?? ExtractImageUrlFromHtml(desc));
                var art = new RssArticle
                {
                    Title = title,
                    Link = ResolveUrl(feedUrl, link),
                    Summary = desc,
                    PublishDate = pub,
                    Source = sourceTitle,
                    ThumbnailUrl = string.IsNullOrWhiteSpace(imgUrl) ? enclosure : imgUrl,
                    EnclosureUrl = enclosure
                };
                return art;
            }
            catch { return null; }
        }

        private RssArticle BuildArticleFromAtomEntry(XElement entry, string sourceTitle, string feedUrl)
        {
            try
            {
                XNamespace ns = entry.GetDefaultNamespace();
                string title = entry.Element(ns + "title")?.Value?.Trim();
                string link = entry.Elements().FirstOrDefault(e => e.Name.LocalName == "link" && (string)e.Attribute("rel") != "self")?.Attribute("href")?.Value;
                string desc = entry.Element(ns + "summary")?.Value ?? entry.Element(ns + "content")?.Value ?? string.Empty;
                DateTime pub = ParseDate(entry.Element(ns + "updated")?.Value) ??
                                ParseDate(entry.Element(ns + "published")?.Value) ??
                                DateTime.UtcNow;
                string enclosure = ResolveUrl(feedUrl, ExtractEnclosureUrl(entry));
                string imgUrl = ResolveUrl(feedUrl, ExtractImageUrl(entry) ?? ExtractImageUrlFromHtml(desc));
                var art = new RssArticle
                {
                    Title = title,
                    Link = ResolveUrl(feedUrl, link),
                    Summary = desc,
                    PublishDate = pub,
                    Source = sourceTitle,
                    ThumbnailUrl = string.IsNullOrWhiteSpace(imgUrl) ? enclosure : imgUrl,
                    EnclosureUrl = enclosure
                };
                return art;
            }
            catch { return null; }
        }

        private void QueueThumbnailForArticle(RssArticle art)
        {
            if (art == null) return;
            _ = Task.Run(async () =>
            {
                try
                {
                    var url = art.ThumbnailUrl;
                    System.Drawing.Image img = await ThumbnailLoader.GetAsync(url, 64).ConfigureAwait(false);

                    // fallback: explicit enclosure, then og:image, then favicon
                    if (img == null && !string.IsNullOrWhiteSpace(art.EnclosureUrl))
                    {
                        img = await ThumbnailLoader.GetAsync(art.EnclosureUrl, 64).ConfigureAwait(false);
                    }
                    if (img == null && !string.IsNullOrWhiteSpace(art.Link))
                    {
                        var og = await FetchOgImageAsync(art.Link, art.Link).ConfigureAwait(false);
                        img = await ThumbnailLoader.GetAsync(og, 64).ConfigureAwait(false);
                        if (img == null)
                        {
                            var fav = await FetchFaviconAsync(art.Link).ConfigureAwait(false);
                            img = await ThumbnailLoader.GetAsync(fav, 32).ConfigureAwait(false);
                        }
                    }

                    if (img != null)
                    {
                        await Dispatcher.InvokeAsync(() => art.ImageSource = ToBitmapImage(img));
                    }
                }
                catch { }
            });
        }
        public BitmapImage ToBitmapImage(System.Drawing.Image bitmap)
        {
            using (var clone = new System.Drawing.Bitmap(bitmap))
            using (var memory = new MemoryStream())
            {
                clone.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bi = new BitmapImage();
                bi.BeginInit();
                bi.StreamSource = memory;
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();
                bi.Freeze();
                return bi;
            }
        }
        private string ExtractEnclosureUrl(XElement parent)
        {
            try
            {
                // RSS <enclosure url="..." type="image/*" />
                var rss = parent.Elements().FirstOrDefault(e => e.Name.LocalName == "enclosure" &&
                    (((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true))?.Attribute("url")?.Value;
                if (!string.IsNullOrWhiteSpace(rss)) return rss;

                // Atom <link rel="enclosure" type="image/*" href="..." />
                var atom = parent.Elements().FirstOrDefault(e => e.Name.LocalName == "link" &&
                    string.Equals((string)e.Attribute("rel"), "enclosure", StringComparison.OrdinalIgnoreCase) &&
                    (((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true))?.Attribute("href")?.Value;
                if (!string.IsNullOrWhiteSpace(atom)) return atom;
            }
            catch { }
            return null;
        }

        // Re-add missing helpers required by the enclosure fallback and list selection
        private static DateTime? ParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (DateTime.TryParse(s, out var dt)) return dt.ToUniversalTime();
            return null;
        }

        private string ExtractImageUrl(XElement parent)
        {
            var enclosure = parent.Elements().FirstOrDefault(e => e.Name.LocalName == "enclosure" &&
                (((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true))?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(enclosure)) return enclosure;

            var thumb = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "thumbnail")?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(thumb)) return thumb;

            var mediaContent = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "content" && (
                (((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true) ||
                string.Equals((string)e.Attribute("medium"), "image", StringComparison.OrdinalIgnoreCase)));
            var mediaUrl = mediaContent?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(mediaUrl)) return mediaUrl;

            var linkEnclosure = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "link" &&
                string.Equals((string)e.Attribute("rel"), "enclosure", StringComparison.OrdinalIgnoreCase) &&
                (((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true));
            var href = linkEnclosure?.Attribute("href")?.Value;
            if (!string.IsNullOrWhiteSpace(href)) return href;

            var itunesImg = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "image" && e.Name.NamespaceName.Contains("itunes"));
            var itHref = itunesImg?.Attribute("href")?.Value ?? itunesImg?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(itHref)) return itHref;

            return null;
        }

        private string ExtractImageUrlFromHtml(string html)
        {
            if (string.IsNullOrWhiteSpace(html)) return null;
            try
            {
                var m = Regex.Match(html, "(?is)<img[^>]+src\\s*=\\s*['\"](?<u>[^'\"]+)['\"]");
                var u = m.Success ? m.Groups["u"].Value : null;
                if (!string.IsNullOrWhiteSpace(u)) return System.Net.WebUtility.HtmlDecode(u);

                var lazyAttrs = new[] { "data-src", "data-original", "data-thumbnail", "data-srcset" };
                foreach (var attr in lazyAttrs)
                {
                    var mm = Regex.Match(html, $"(?is)<img[^>]+{attr}\\s*=\\s*['\"](?<u>[^'\"]+)['\"]");
                    if (mm.Success)
                    {
                        var val = mm.Groups["u"].Value;
                        if (!string.IsNullOrWhiteSpace(val)) return System.Net.WebUtility.HtmlDecode(val.Split(' ').FirstOrDefault());
                    }
                }

                var mset = Regex.Match(html, "(?is)<img[^>]+srcset\\s*=\\s*['\"](?<s>[^'\"]+)['\"]");
                if (mset.Success)
                {
                    var list = mset.Groups["s"].Value;
                    var first = list.Split(',').Select(p => p.Trim().Split(' ')[0]).FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(first)) return System.Net.WebUtility.HtmlDecode(first);
                }
            }
            catch { }
            return null;
        }

        private static string ResolveUrl(string baseUrl, string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return url;
            try
            {
                url = System.Net.WebUtility.HtmlDecode(url);
                if (url.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return url;
                if (url.StartsWith("//"))
                {
                    var scheme = new Uri(baseUrl).Scheme;
                    return scheme + ":" + url;
                }
                if (Uri.IsWellFormedUriString(url, UriKind.Absolute)) return url;
                var b = new Uri(baseUrl);
                var abs = new Uri(b, url);
                return abs.AbsoluteUri;
            }
            catch { return url; }
        }

        private async Task<string> FetchOgImageAsync(string url, string baseUrl)
        {
            try
            {
                using (var wc = new WebClient())
                {
                    wc.Headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36";
                    var html = await wc.DownloadStringTaskAsync(url);
                    string og = null;
                    var m1 = Regex.Match(html, "(?is)<meta[^>]+property=['\"]og:image['\"][^>]+content=['\"](?<u>[^'\"]+)['\"]");
                    if (m1.Success) og = m1.Groups["u"].Value;
                    if (string.IsNullOrWhiteSpace(og))
                    {
                        var m2 = Regex.Match(html, "(?is)<meta[^>]+name=['\"]twitter:image['\"][^>]+content=['\"](?<u>[^'\"]+)['\"]");
                        if (m2.Success) og = m2.Groups["u"].Value;
                    }
                    og = ResolveUrl(url, og);
                    return og;
                }
            }
            catch { return null; }
        }

        private async Task<string> FetchFaviconAsync(string pageUrl)
        {
            if (string.IsNullOrWhiteSpace(pageUrl)) return null;
            try
            {
                var uri = new Uri(pageUrl);
                using (var wc = new WebClient())
                {
                    wc.Headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)";
                    string html = null;
                    try { html = await wc.DownloadStringTaskAsync(pageUrl); } catch { html = null; }
                    if (!string.IsNullOrEmpty(html))
                    {
                        var m = Regex.Match(html, "(?is)<link[^>]+rel=['\"](?:shortcut )?icon['\"][^>]+href=['\"](?<u>[^'\"]+)['\"]");
                        if (m.Success)
                        {
                            var u = m.Groups["u"].Value;
                            return ResolveUrl(pageUrl, System.Net.WebUtility.HtmlDecode(u));
                        }
                    }
                    var fallback = uri.Scheme + "://" + uri.Host + "/favicon.ico";
                    return fallback;
                }
            }
            catch { return null; }
        }

        private Microsoft.Web.WebView2.Wpf.WebView2 GetWv() => this.FindName("wbArticle") as Microsoft.Web.WebView2.Wpf.WebView2;

        private void RenderArticleHtml(RssArticle art)
        {
            var wv = GetWv();
            if (wv?.CoreWebView2 == null)
            {
                _ = Dispatcher.InvokeAsync(async () => { await InitArticleWebViewAsync(); RenderArticleHtml(art); });
                return;
            }
            var title = System.Net.WebUtility.HtmlEncode(art.Title ?? string.Empty);
            var meta = System.Net.WebUtility.HtmlEncode(string.Format("{0} | {1:dd.MM.yyyy HH:mm}", art.Source, art.PublishDate));
            var content = NormalizeArticleHtml(art.Summary ?? string.Empty, art.Link);
            var html = $@"<!doctype html><html><head><meta charset='utf-8'>
<base href='{System.Net.WebUtility.HtmlEncode(art.Link ?? "about:blank")}' />
<style>body{{font-family:Segoe UI,Arial,Helvetica,sans-serif;margin:0;padding:12px;line-height:1.4}} img{{max-width:100%;height:auto}} a{{color:#0a84ff;text-decoration:none}}</style>
</head><body>
<h2>{title}</h2>
<div style='color:#666;font-size:12px;margin-bottom:12px'>{meta}</div>
<div>{content}</div>
</body></html>";
            try { wv.NavigateToString(html); } catch { }
        }

        private string NormalizeArticleHtml(string html, string baseUrl)
        {
            try
            {
                html = Regex.Replace(html, "(?is)<script.*?>.*?</script>", string.Empty);
                html = Regex.Replace(html, "(?is)src=['\"]//", m => "src='" + new Uri(baseUrl ?? "https://", UriKind.Absolute).Scheme + "://");
                html = Regex.Replace(html, "(?is)<img([^>]+)(data-src|data-original)=['\"](?<u>[^'\"]+)['\"]", m =>
                {
                    var u = ResolveUrl(baseUrl ?? "about:blank", System.Net.WebUtility.HtmlDecode(m.Groups["u"].Value));
                    return "<img" + m.Groups[1].Value + "src='" + u + "'";
                });
                html = Regex.Replace(html, "(?is)srcset=['\"](?<s>[^'\"]+)['\"]", m =>
                {
                    var first = m.Groups["s"].Value.Split(',').Select(p => p.Trim().Split(' ')[0]).FirstOrDefault();
                    if (string.IsNullOrWhiteSpace(first)) return string.Empty;
                    var u = ResolveUrl(baseUrl ?? "about:blank", System.Net.WebUtility.HtmlDecode(first));
                    return "src='" + u + "'";
                });
                return html;
            }
            catch { return html; }
        }

        private void lstArticles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstArticles.SelectedItem is RssArticle art)
            {
                txtArticleTitle.Text = art.Title;
                txtArticleMeta.Text = string.Format("{0} | {1:dd.MM.yyyy HH:mm}", art.Source, art.PublishDate);
                RenderArticleHtml(art);
            }
        }

        private static class ThumbnailLoader
        {
            private static readonly ConcurrentDictionary<string, System.Drawing.Image> _mem = new ConcurrentDictionary<string, System.Drawing.Image>(StringComparer.OrdinalIgnoreCase);
            private static readonly string _diskDir = InitDisk();

            private static string InitDisk()
            {
                try
                {
                    var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VivitControlCenter", "Cache", "thumbnails");
                    Directory.CreateDirectory(dir);
                    return dir;
                }
                catch { return null; }
            }

            public static async Task<System.Drawing.Image> GetAsync(string url, int decodeWidth)
            {
                if (string.IsNullOrWhiteSpace(url)) return null;

                // Pfad früh berechnen, damit wir auch data:-URLs speichern können
                var diskFile = SafePath(url);

                if (url.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var comma = url.IndexOf(',');
                        if (comma > 0)
                        {
                            var meta = url.Substring(5, comma - 5);
                            var payload = url.Substring(comma + 1);
                            byte[] bytes = meta.ToLowerInvariant().Contains(";base64")
                                ? Convert.FromBase64String(payload)
                                : System.Text.Encoding.UTF8.GetBytes(Uri.UnescapeDataString(payload));

                            var bmp = BuildBitmap(bytes, decodeWidth);
                            if (bmp != null)
                            {
                                // in Memory-Cache aufnehmen
                                _mem[url] = bmp;
                                // auf Platte persistieren
                                TryWriteFile(diskFile, bytes);
                            }
                            return bmp;
                        }
                    }
                    catch { return null; }
                }

                if (_mem.TryGetValue(url, out var cached)) return cached;

                try
                {
                    if (!string.IsNullOrEmpty(diskFile) && File.Exists(diskFile))
                    {
                        byte[] bytes;
                        using (var fs = new FileStream(diskFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            bytes = new byte[fs.Length];
                            await fs.ReadAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
                        }
                        var bmp = BuildBitmap(bytes, decodeWidth);
                        if (bmp != null)
                        {
                            _mem[url] = bmp;
                            return bmp;
                        }
                        // optional: defekte Datei entfernen
                        try { File.Delete(diskFile); } catch { }
                    }
                }
                catch { 

                }

                //await _thumbConcurrency.WaitAsync().ConfigureAwait(false);
                try
                {
                    using (var resp = await _imageHttp.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                    {
                        if (!resp.IsSuccessStatusCode) return null;

                        var ct = resp.Content.Headers.ContentType?.MediaType ?? string.Empty;
                        if (ct.IndexOf("webp", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            ct.IndexOf("avif", StringComparison.OrdinalIgnoreCase) >= 0)
                            return null;

                        var bytes = await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                        if (bytes == null || bytes.Length == 0) return null;
                        if (LooksLikeWebpOrAvif(bytes)) return null;

                        var bmp = BuildBitmap(bytes, decodeWidth);
                        if (bmp != null)
                        {
                            _mem[url] = bmp;
                            TryWriteFile(diskFile, bytes);
                        }
                        return bmp;
                    }
                }
                catch (Exception e){ 
                    return null; 
                }
                finally { _thumbConcurrency.Release(); }
            }

            private static bool LooksLikeWebpOrAvif(byte[] data)
            {
                try
                {
                    if (data.Length > 12)
                    {
                        if (data[0] == 0x52 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x46 && data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50)
                            return true;
                        var header = System.Text.Encoding.ASCII.GetString(data, 4, Math.Min(12, data.Length - 4));
                        if (header.IndexOf("ftypavif", StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
                catch { }
                return false;
            }

            private static void TryWriteFile(string path, byte[] data)
            {
                try
                {
                    if (string.IsNullOrEmpty(path) || data == null || data.Length == 0) return;
                    File.WriteAllBytes(path, data);
                }
                catch { }
            }

            private static string SafePath(string url)
            {
                try
                {
                    if (string.IsNullOrEmpty(_diskDir)) return null;
                    using (var sha1 = SHA1.Create())
                    {
                        var hash = sha1.ComputeHash(System.Text.Encoding.UTF8.GetBytes(url));
                        var name = BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant() + ".bin";
                        return Path.Combine(_diskDir, name);
                    }
                }
                catch { return null; }
            }

            private static System.Drawing.Image BuildBitmap(byte[] data, int decodeWidth)
            {
                try
                {
                    using (var ms = new MemoryStream(data))
                    using (var img = System.Drawing.Image.FromStream(ms, useEmbeddedColorManagement: false, validateImageData: true))
                    {
                        // optionales Downscaling
                        if (decodeWidth > 0 && img.Width > decodeWidth)
                        {
                            int w = decodeWidth;
                            int h = (int)Math.Round((double)img.Height * w / img.Width);
                            var bmp = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                            using (var g = System.Drawing.Graphics.FromImage(bmp))
                            {
                                g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                                g.DrawImage(img, 0, 0, w, h);
                            }
                            return bmp; // vollständig vom Stream entkoppelt
                        }

                        return new System.Drawing.Bitmap(img); // entkoppelte Kopie
                    }
                }
                catch { return null; }
            }
        }
    }

    public class RssArticle : INotifyPropertyChanged
    {
        public string Title { get; set; }
        public DateTime PublishDate { get; set; }
        public string Source { get; set; }
        public string Link { get; set; }
        public string Summary { get; set; }

        public string ThumbnailUrl { get; set; }
        public string EnclosureUrl { get; set; }

        private ImageSource _imageSource;
        public ImageSource ImageSource
        {
            get => _imageSource;
            set { if (!Equals(_imageSource, value)) { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); } }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}