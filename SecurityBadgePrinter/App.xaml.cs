using System;
using System.IO;
using System.Windows;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using SecurityBadgePrinter.Models;
using SecurityBadgePrinter.Services;

namespace SecurityBadgePrinter
{
    public partial class App : Application
    {
        public static IConfiguration Configuration { get; private set; } = default!;
        public static AppConfig Config { get; private set; } = default!;
        public static GraphServiceClient GraphClient { get; private set; } = default!;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var builder = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            Configuration = builder.Build();

            Config = Configuration.Get<AppConfig>() ?? new AppConfig();

            GraphClient = GraphServiceFactory.BuildGraphClient(Config);
        }
    }
}
