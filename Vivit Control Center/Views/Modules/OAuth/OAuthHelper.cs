using System;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules.OAuth
{
    public static class OAuthHelper
    {
        public class OAuthTokenResult
        {
            public string AccessToken { get; set; }
            public string RefreshToken { get; set; }
            public DateTime ExpiryUtc { get; set; }
        }

        public static async Task<OAuthTokenResult> AcquireTokensInteractiveAsync(EmailAccount acc)
        {
            var provider = (acc.OAuthProvider ?? string.Empty).Trim().ToLowerInvariant();
            var tenant = string.IsNullOrWhiteSpace(acc.OAuthTenant) ? "common" : acc.OAuthTenant.Trim();
            string authEndpoint;
            string tokenEndpoint;
            string scope = acc.OAuthScope;

            if (!string.IsNullOrWhiteSpace(acc.OAuthTokenEndpoint))
            {
                tokenEndpoint = acc.OAuthTokenEndpoint;
            }
            else if (provider == "google")
            {
                tokenEndpoint = "https://oauth2.googleapis.com/token";
            }
            else if (provider == "microsoft")
            {
                tokenEndpoint = $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token";
            }
            else
            {
                tokenEndpoint = null; // required for Custom via settings
            }

            if (provider == "google")
            {
                authEndpoint = "https://accounts.google.com/o/oauth2/v2/auth";
                if (string.IsNullOrWhiteSpace(scope)) scope = "https://mail.google.com/"; // Gmail IMAP/SMTP
            }
            else if (provider == "microsoft")
            {
                authEndpoint = $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize";
                if (string.IsNullOrWhiteSpace(scope)) scope = "offline_access https://outlook.office365.com/IMAP.AccessAsUser.All https://outlook.office365.com/SMTP.Send";
            }
            else
            {
                // Custom
                if (string.IsNullOrWhiteSpace(acc.OAuthTokenEndpoint)) throw new InvalidOperationException("Provide Token Endpoint for Custom OAuth provider.");
                if (string.IsNullOrWhiteSpace(scope)) throw new InvalidOperationException("Provide Scope(s) for Custom OAuth provider.");
                if (string.IsNullOrWhiteSpace(acc.OAuthClientId)) throw new InvalidOperationException("Provide Client Id for Custom OAuth provider.");
                authEndpoint = ResolveCustomAuthorizeEndpoint(tokenEndpoint); // attempt same base
            }

            // Loopback redirect
            var (listener, redirectUri) = CreateLoopbackListener();

            // PKCE
            var verifier = CreatePkceCodeVerifier();
            var challenge = CreatePkceCodeChallenge(verifier);

            // Build auth URL
            var state = Guid.NewGuid().ToString("N");
            var url = new StringBuilder();
            url.Append(authEndpoint);
            url.Append("?response_type=code");
            url.Append("&client_id=").Append(Uri.EscapeDataString(acc.OAuthClientId ?? string.Empty));
            url.Append("&redirect_uri=").Append(Uri.EscapeDataString(redirectUri));
            url.Append("&scope=").Append(Uri.EscapeDataString(scope));
            url.Append("&state=").Append(Uri.EscapeDataString(state));
            url.Append("&code_challenge=").Append(Uri.EscapeDataString(challenge));
            url.Append("&code_challenge_method=S256");
            if (provider == "google")
            {
                // Request refresh token
                url.Append("&access_type=offline&prompt=consent");
            }

            // Launch in system browser
            try { Process.Start(new ProcessStartInfo(url.ToString()) { UseShellExecute = true }); } catch { }

            // Wait for redirect
            var ctx = await listener.GetContextAsync();
            string responseHtml = "<html><body>You may close this window.</body></html>";
            try
            {
                var req = ctx.Request;
                var code = req.QueryString["code"]; var receivedState = req.QueryString["state"]; var error = req.QueryString["error"];
                if (!string.IsNullOrEmpty(error)) throw new Exception("OAuth error: " + error);
                if (string.IsNullOrEmpty(code)) throw new Exception("No authorization code.");
                if (!string.Equals(state, receivedState)) throw new Exception("State mismatch.");

                var token = await ExchangeCodeForTokenAsync(tokenEndpoint, acc, code, redirectUri, verifier, scope);
                responseHtml = "<html><body>Authentication complete. You can return to the app.</body></html>";
                return token;
            }
            finally
            {
                try
                {
                    var bytes = Encoding.UTF8.GetBytes(responseHtml);
                    ctx.Response.ContentType = "text/html";
                    ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
                    ctx.Response.OutputStream.Close();
                    listener.Stop();
                    listener.Close();
                }
                catch { }
            }
        }

        private static (HttpListener listener, string redirectUri) CreateLoopbackListener()
        {
            // Bind to 127.0.0.1 on a free port
            int port = GetFreeTcpPort();
            string redirect = $"http://127.0.0.1:3000/oauth2redirect";
            var listener = new HttpListener();
            listener.Prefixes.Add($"http://127.0.0.1:3000/");
            listener.Start();
            return (listener, redirect);
        }

        private static int GetFreeTcpPort()
        {
            var l = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
            l.Start();
            int port = ((System.Net.IPEndPoint)l.LocalEndpoint).Port;
            l.Stop();
            return port;
        }

        private static string CreatePkceCodeVerifier()
        {
            var bytes = new byte[32];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(bytes);
            return Base64UrlEncode(bytes);
        }
        private static string CreatePkceCodeChallenge(string verifier)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.ASCII.GetBytes(verifier));
                return Base64UrlEncode(bytes);
            }
        }
        private static string Base64UrlEncode(byte[] arg)
        {
            string s = Convert.ToBase64String(arg);
            s = s.Replace('+', '-').Replace('/', '_').TrimEnd('=');
            return s;
        }

        private static async Task<OAuthTokenResult> ExchangeCodeForTokenAsync(string tokenEndpoint, EmailAccount acc, string code, string redirectUri, string codeVerifier, string scope)
        {
            if (string.IsNullOrWhiteSpace(tokenEndpoint)) throw new InvalidOperationException("Token endpoint is required");
            using (var http = new HttpClient())
            {
                var body = new StringBuilder();
                body.Append("grant_type=authorization_code");
                body.Append("&code=").Append(Uri.EscapeDataString(code));
                body.Append("&redirect_uri=").Append(Uri.EscapeDataString(redirectUri));
                body.Append("&client_id=").Append(Uri.EscapeDataString(acc.OAuthClientId ?? string.Empty));
                if (!string.IsNullOrEmpty(acc.OAuthClientSecret))
                    body.Append("&client_secret=").Append(Uri.EscapeDataString(acc.OAuthClientSecret));
                body.Append("&code_verifier=").Append(Uri.EscapeDataString(codeVerifier));
                if (!string.IsNullOrWhiteSpace(scope))
                    body.Append("&scope=").Append(Uri.EscapeDataString(scope));

                var content = new StringContent(body.ToString(), Encoding.UTF8, "application/x-www-form-urlencoded");
                var resp = await http.PostAsync(tokenEndpoint, content);
                var respText = await resp.Content.ReadAsStringAsync();
                if (!resp.IsSuccessStatusCode) throw new Exception($"Token exchange failed: {resp.StatusCode} {respText}");

                var access = Regex.Match(respText, "\\\"access_token\\\"\\s*:\\s*\\\"(.*?)\\\"").Groups[1].Value;
                var expires = Regex.Match(respText, "\\\"expires_in\\\"\\s*:\\s*(\\d+)").Groups[1].Value;
                var refresh = Regex.Match(respText, "\\\"refresh_token\\\"\\s*:\\s*\\\"(.*?)\\\"").Groups[1].Value;
                int expiresSec = 3600; int.TryParse(expires, out expiresSec);
                return new OAuthTokenResult
                {
                    AccessToken = access,
                    RefreshToken = string.IsNullOrEmpty(refresh) ? acc.OAuthRefreshToken : refresh,
                    ExpiryUtc = DateTime.UtcNow.AddSeconds(expiresSec - 60)
                };
            }
        }

        private static string ResolveCustomAuthorizeEndpoint(string tokenEndpoint)
        {
            // Best-effort guess: replace /token with /authorize
            if (string.IsNullOrEmpty(tokenEndpoint)) return null;
            if (tokenEndpoint.Contains("/token")) return tokenEndpoint.Replace("/token", "/authorize");
            return tokenEndpoint;
        }
    }
}
