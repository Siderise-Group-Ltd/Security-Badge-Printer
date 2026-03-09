using System;
using System.IO;
using Azure.Identity;
using Azure.Core;
using Microsoft.Graph;
using SecurityBadgePrinter.Models;

namespace SecurityBadgePrinter.Services
{
    /// <summary>
    /// Factory for creating Microsoft Graph clients for the Siderise Security Badge Printer
    /// </summary>
    public static class GraphServiceFactory
    {
        private static readonly string[] Scopes = new[] { "User.Read", "User.Read.All", "Group.Read.All" };

        public static GraphServiceClient BuildGraphClient(AppConfig config)
        {
            var authDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SecurityBadgePrinter");
            var authRecordPath = Path.Combine(authDir, "authRecord.json");

            AuthenticationRecord? record = null;
            try
            {
                if (File.Exists(authRecordPath))
                {
                    using var fs = File.OpenRead(authRecordPath);
                    record = AuthenticationRecord.Deserialize(fs);
                }
            }
            catch { }

            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = config.AzureAd.TenantId,
                ClientId = config.AzureAd.ClientId,
                RedirectUri = new Uri("http://localhost"),
                LoginHint = "SEC@siderise.com",
                TokenCachePersistenceOptions = new TokenCachePersistenceOptions
                {
                    Name = "SecurityBadgePrinter",
                    UnsafeAllowUnencryptedStorage = true
                },
                AuthenticationRecord = record
            };

            var credential = new InteractiveBrowserCredential(options);

            var client = new GraphServiceClient(credential, Scopes);

            // If no record yet, authenticate once to capture it, then persist
            if (record == null)
            {
                try
                {
                    var newRecord = credential.Authenticate(new TokenRequestContext(Scopes));
                    Directory.CreateDirectory(authDir);
                    using var fs = File.Create(authRecordPath);
                    newRecord.Serialize(fs);
                }
                catch { /* Will fallback to interactive on first Graph call */ }
            }
            return client;
        }
    }
}
