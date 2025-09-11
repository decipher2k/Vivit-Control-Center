using Microsoft.Web.WebView2.Core;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class BrowserModule : BaseSimpleModule
    {
        public BrowserModule()
        {
            InitializeComponent();
            InitializeAsync();
        }

        private async void InitializeAsync()
        {
            try
            {
                await webView.EnsureCoreWebView2Async();
                SignalLoadedOnce();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WebView2 konnte nicht initialisiert werden: {ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Navigate(string url)
        {
            try
            {
                if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                {
                    url = "https://" + url;
                }
                webView.CoreWebView2.Navigate(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei der Navigation: {ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            if (webView.CanGoBack)
                webView.GoBack();
        }

        private void btnForward_Click(object sender, RoutedEventArgs e)
        {
            if (webView.CanGoForward)
                webView.GoForward();
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            webView.Reload();
        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            Navigate(txtUrl.Text);
        }

        private void txtUrl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Navigate(txtUrl.Text);
                e.Handled = true;
            }
        }

        private void webView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            txtUrl.Text = webView.Source.AbsoluteUri;

            // Update navigation buttons
            btnBack.IsEnabled = webView.CanGoBack;
            btnForward.IsEnabled = webView.CanGoForward;
        }
    }
}