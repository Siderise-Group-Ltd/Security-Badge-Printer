using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using SkiaSharp;

namespace SecurityBadgePrinter.Services
{
    /// <summary>
    /// Handles printing of Siderise security badges to the Zebra ZC300 card printer
    /// </summary>
    public class PrinterService
    {
        private readonly string _printerName;
        private byte[]? _imageBytes;

        public PrinterService(string printerName)
        {
            _printerName = printerName;
        }

        public void Print(SKBitmap badge)
        {
            _imageBytes = BadgeRenderer.EncodeToPng(badge);

            using var printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = _printerName;

            // Landscape print for horizontal card
            printDoc.DefaultPageSettings.Landscape = true;

            // Zero margins so we target the full card area
            printDoc.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            printDoc.OriginAtMargins = false;

            // Try to pick the printer's native card size; fallback to custom CR80 (hundredths of an inch)
            TrySetCardPaperSize(printDoc);

            // Optional: use standard print controller to avoid any UI
            printDoc.PrintController = new StandardPrintController();

            printDoc.PrintPage += OnPrintPage;
            printDoc.Print();
        }

        private static void TrySetCardPaperSize(PrintDocument printDoc)
        {
            try
            {
                // Force a fixed CR80 physical size (landscape) regardless of driver-exposed names
                var custom = new PaperSize("CR80 (custom)", 337, 213)
                {
                    RawKind = (int)PaperKind.Custom
                };
                printDoc.DefaultPageSettings.PaperSize = custom;
            }
            catch
            {
                // Ignore issues; driver will fall back to its default; scaling code still runs
            }
        }

        private void OnPrintPage(object? sender, PrintPageEventArgs e)
        {
            if (_imageBytes == null) return;
            if (e == null || e.Graphics == null) return;

            using var ms = new MemoryStream(_imageBytes);
            using var img = Image.FromStream(ms);
            if (img == null) { e.HasMorePages = false; return; }

            // Work in Display units (1/100 inch) to avoid anisotropic DPI distortions
            float hardX = e.PageSettings.HardMarginX; // 1/100 inch
            float hardY = e.PageSettings.HardMarginY; // 1/100 inch

            // Improve quality
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

            // Edge-to-edge: map physical page to device pixels (accounts for different DpiX/DpiY)
            var page = e.PageBounds; // in 1/100 inch
            float dpix = e.Graphics.DpiX;
            float dpiy = e.Graphics.DpiY;

            // Convert hard margins and page size to pixels
            float hardXpx = hardX / 100f * dpix;
            float hardYpx = hardY / 100f * dpiy;
            int pageWpx = (int)Math.Round(page.Width / 100f * dpix);
            int pageHpx = (int)Math.Round(page.Height / 100f * dpiy);
            if (pageWpx < pageHpx)
            {
                var tmp = pageWpx; pageWpx = pageHpx; pageHpx = tmp;
            }

            // Diagnostics
            try
            {
                Serilog.Log.Information("Print geom: DpiX={DpiX} DpiY={DpiY} Page(1/100in)={W}x{H} HardMargins(1/100in)={HX},{HY} => Page(px)={PW}x{PH}", dpix, dpiy, page.Width, page.Height, hardX, hardY, pageWpx, pageHpx);
            }
            catch { }

            // Draw in pixels using exact physical CR80 size from 300 DPI design
            e.Graphics.PageUnit = System.Drawing.GraphicsUnit.Pixel;
            e.Graphics.TranslateTransform(-hardXpx, -hardYpx);

            // Our badge is designed at 300 DPI (Width/Height in pixels in BadgeRenderer)
            int targetWpx = (int)Math.Round(BadgeRenderer.Width * (dpix / 300f));
            int targetHpx = (int)Math.Round(BadgeRenderer.Height * (dpiy / 300f));

            // Center the target within the physical page in pixels
            int offsetXpx = (pageWpx - targetWpx) / 2;
            int offsetYpx = (pageHpx - targetHpx) / 2;

            // If page is smaller than target due to driver constraints, clamp to page and preserve aspect
            if (targetWpx > pageWpx || targetHpx > pageHpx)
            {
                double scale = Math.Min(pageWpx / (double)targetWpx, pageHpx / (double)targetHpx);
                targetWpx = (int)Math.Round(targetWpx * scale);
                targetHpx = (int)Math.Round(targetHpx * scale);
                offsetXpx = (pageWpx - targetWpx) / 2;
                offsetYpx = (pageHpx - targetHpx) / 2;
            }

            var src = new RectangleF(0, 0, img.Width, img.Height);
            var dest = new Rectangle(offsetXpx, offsetYpx, targetWpx, targetHpx);
            e.Graphics.DrawImage(img, dest, src, System.Drawing.GraphicsUnit.Pixel);
            e.HasMorePages = false;
        }
    }
}
