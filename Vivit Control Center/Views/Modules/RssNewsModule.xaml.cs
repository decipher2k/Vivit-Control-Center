using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using System.Xml;
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

        public RssNewsModule()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += async (_, __) => await EnsureLoadedAsync();
        }

        private async Task EnsureLoadedAsync()
        {
            if (_loaded) return;
            _loaded = true;
            _settings = AppSettings.Load();
            await LoadFeedsAsync();
        }

        public override async Task PreloadAsync()
        {
            await EnsureLoadedAsync();
            await base.PreloadAsync();
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

            // Alle Feeds einlesen (jedes sortiert nach Datum DESC, lokal begrenzt auf totalTarget zur Schonung)
            var perFeedLists = new System.Collections.Generic.List<System.Collections.Generic.List<RssArticle>>();
            foreach (var feed in feeds)
            {
                var list = await FetchFeedArticlesAsync(feed, totalTarget);
                perFeedLists.Add(list);
            }

            // Round-Robin über die per-Feed-Listen bis das Ziel erreicht ist
            var final = new System.Collections.Generic.List<RssArticle>(totalTarget);
            bool progress = true;
            while (final.Count < totalTarget && progress)
            {
                progress = false;
                for (int i = 0; i < perFeedLists.Count && final.Count < totalTarget; i++)
                {
                    var list = perFeedLists[i];
                    if (list.Count == 0) continue;
                    final.Add(list[0]);
                    list.RemoveAt(0);
                    progress = true;
                }
            }

            // CollectionView mit Gruppierung (nach Source)
            var view = new ListCollectionView(final);
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(RssArticle.Source)));
            lstArticles.ItemsSource = view;
        }

        private async Task<System.Collections.Generic.List<RssArticle>> FetchFeedArticlesAsync(string feed, int softLimit)
        {
            var result = new System.Collections.Generic.List<RssArticle>();
            try
            {
                var bytes = await _http.GetByteArrayAsync(feed);
                var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(bytes));
                if (xml.Root == null) return result;

                if (xml.Root.Name.LocalName.Equals("rss", StringComparison.OrdinalIgnoreCase) ||
                    xml.Root.Elements().Any(e => e.Name.LocalName == "channel"))
                {
                    var channel = xml.Root.Element("channel") ?? xml.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "channel");
                    var sourceTitle = channel?.Element("title")?.Value?.Trim() ?? feed;
                    foreach (var item in channel?.Elements().Where(e => e.Name.LocalName == "item") ?? Enumerable.Empty<XElement>())
                    {
                        var art = BuildArticleFromRssItem(item, sourceTitle);
                        if (art != null) result.Add(art);
                        if (result.Count >= softLimit) break; // lokal begrenzen
                    }
                }
                else if (xml.Root.Name.LocalName.Equals("feed", StringComparison.OrdinalIgnoreCase))
                {
                    var sourceTitle = xml.Root.Element(xml.Root.GetDefaultNamespace() + "title")?.Value?.Trim() ?? feed;
                    foreach (var entry in xml.Root.Elements().Where(e => e.Name.LocalName == "entry"))
                    {
                        var art = BuildArticleFromAtomEntry(entry, sourceTitle);
                        if (art != null) result.Add(art);
                        if (result.Count >= softLimit) break;
                    }
                }
            }
            catch { }

            return result.OrderByDescending(a => a.PublishDate).ToList();
        }

        private RssArticle BuildArticleFromRssItem(XElement item, string sourceTitle)
        {
            try
            {
                string title = item.Element("title")?.Value?.Trim();
                string link = item.Element("link")?.Value?.Trim();
                string desc = item.Element("description")?.Value ?? item.Element("content")?.Value ?? string.Empty;
                DateTime pub = ParseDate(item.Element("pubDate")?.Value) ?? DateTime.UtcNow;
                string imgUrl = ExtractImageUrl(item);
                var art = new RssArticle
                {
                    Title = title,
                    Link = link,
                    Summary = desc,
                    PublishDate = pub,
                    Source = sourceTitle
                };
                if (!string.IsNullOrWhiteSpace(imgUrl)) art.Image = TryLoadBitmap(imgUrl);
                return art;
            }
            catch { return null; }
        }

        private RssArticle BuildArticleFromAtomEntry(XElement entry, string sourceTitle)
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
                string imgUrl = ExtractImageUrl(entry);
                var art = new RssArticle
                {
                    Title = title,
                    Link = link,
                    Summary = desc,
                    PublishDate = pub,
                    Source = sourceTitle
                };
                if (!string.IsNullOrWhiteSpace(imgUrl)) art.Image = TryLoadBitmap(imgUrl);
                return art;
            }
            catch { return null; }
        }

        private static DateTime? ParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (DateTime.TryParse(s, out var dt)) return dt.ToUniversalTime();
            return null;
        }

        private string ExtractImageUrl(XElement parent)
        {
            var enclosure = parent.Elements().FirstOrDefault(e => e.Name.LocalName == "enclosure" &&
                ((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true)?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(enclosure)) return enclosure;
            var thumb = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "thumbnail")?.Attribute("url")?.Value;
            if (!string.IsNullOrWhiteSpace(thumb)) return thumb;
            var contentImg = parent.Descendants().FirstOrDefault(e => e.Name.LocalName == "content" &&
                ((string)e.Attribute("type"))?.StartsWith("image", StringComparison.OrdinalIgnoreCase) == true)?.Attribute("url")?.Value;
            return contentImg;
        }

        private BitmapImage TryLoadBitmap(string url)
        {
            try
            {
                var data = _http.GetByteArrayAsync(url).GetAwaiter().GetResult();
                using (var ms = new System.IO.MemoryStream(data))
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
            }
            catch { return null; }
        }

        private void lstArticles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstArticles.SelectedItem is RssArticle art)
            {
                txtArticleTitle.Text = art.Title;
                txtArticleMeta.Text = $"{art.Source} | {art.PublishDate:dd.MM.yyyy HH:mm}";
                txtArticleContent.Text = StripHtml(art.Summary);
                if (art.Image != null)
                {
                    imgArticle.Source = art.Image;
                    imgArticle.Visibility = Visibility.Visible;
                }
                else
                {
                    imgArticle.Visibility = Visibility.Collapsed;
                }
            }
        }

        private static string StripHtml(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            try { return System.Text.RegularExpressions.Regex.Replace(input, "<.*?>", string.Empty).Trim(); }
            catch { return input; }
        }
    }

    public class RssArticle
    {
        public string Title { get; set; }
        public DateTime PublishDate { get; set; }
        public string Source { get; set; }
        public string Link { get; set; }
        public string Summary { get; set; }
        public BitmapImage Image { get; set; }
    }
}