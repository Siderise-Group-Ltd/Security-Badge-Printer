using System.Collections.Generic;
 
namespace SecurityBadgePrinter.Models
{
    public class AppConfig
    {
        public AzureAdConfig AzureAd { get; set; } = new AzureAdConfig();
        public PrinterConfig Printer { get; set; } = new PrinterConfig();
    }

    public class AzureAdConfig
    {
        public string TenantId { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public List<string> AllowedGroupIds { get; set; } = new List<string>();
    }

    public class PrinterConfig
    {
        public string Name { get; set; } = "Zebra - Security Badge";
    }
}
