using System;
using System.IO;
using System.Reflection;
using QRCoder;
using SkiaSharp;

namespace SecurityBadgePrinter.Services
{
    /// <summary>
    /// Renders Siderise branded security badges with company colors, typography, and logo
    /// following Siderise Brand Guidelines 2025
    /// </summary>
    public class BadgeRenderer
    {
        // CR80 at 300 DPI (3.375" x 2.125") - Standard security badge size
        public const int Width = 1011;  // 3.37 * 300 ~ 1011
        public const int Height = 638;  // 2.125 * 300 ~ 638
        private const float Dpi = 300f;
        private const float PhotoBorderStroke = 8f; // px thickness of the photo border

        public SKBitmap Render(string displayName, string jobTitle, string department, string upn, SKBitmap? photo)
        {
            var badge = new SKBitmap(Width, Height);
            using var canvas = new SKCanvas(badge);
            canvas.Clear(SKColors.White);

            // Layout margins
            // Top/bottom: 4mm unprintable margin so content is tight to the printable edge
            float topMargin = MmToPx(4f);
            float bottomMargin = MmToPx(4f);
            // Left/right: keep visual spacing similar to before
            float leftMargin = 20f;
            float rightMargin = 20f;
            // Siderise Brand colors from Brand Guidelines 2025
            var brandNavy = new SKColor(0x00, 0x27, 0x51);   // #002751 Navy Primary - for name, department, QR
            var brandBlue = new SKColor(0x00, 0x90, 0xD0);   // #0090D0 Blue Primary - for role/title

            // Left photo panel - matching mockup proportions
            int photoPanelWidth = 460; // Larger photo area like in mockup
            var photoRect = new SKRect(leftMargin, topMargin, leftMargin + photoPanelWidth, Height - bottomMargin);
            DrawPhoto(canvas, photoRect, photo, displayName);

            // Right panel layout - matching mockup
            float rightX = photoRect.Right + 40; // More spacing from photo
            float rightW = Width - rightMargin - rightX;
            float cursorY = topMargin;

            // Top brand logo image - positioned like mockup (top right)
            var logo = LoadLogoBitmap();
            if (logo != null)
            {
                float targetLogoH = 75f; // larger logo for better text readability
                float scale = targetLogoH / logo.Height;
                float logoW = logo.Width * scale;
                float logoH = targetLogoH;
                float logoX = (float)Math.Round(rightX + (rightW - logoW));
                float logoY = (float)Math.Round(cursorY + 10);
                using var logoPaint = new SKPaint { IsAntialias = true, FilterQuality = SKFilterQuality.High };
                canvas.DrawBitmap(logo, new SKRect(logoX, logoY, logoX + logoW, logoY + logoH), logoPaint);
            }
            // Generate QR for bottom right positioning like mockup
            int qrSize = 140; // Smaller QR like in mockup
            string qrContent = BuildQr(upn);
            using var qr = GenerateQrBitmap(qrContent, qrSize, brandNavy);
            
            // QR positioning: bottom-right corner like mockup
            int qrX = (int)(rightX + rightW - qr.Width);
            int qrY = (int)(Height - bottomMargin - qr.Height);
            var qrRect = new SKRect(qrX, qrY, qrX + qr.Width, qrY + qr.Height);

            // Calculate available vertical space for text distribution
            float availableTextHeight = qrY - (topMargin + 80) - 20; // Space between logo area and QR, minus buffer
            
            // Start text positioning from top
            cursorY = topMargin + 115; // More space for larger logo area

            // Name - large and prominent, using full width initially
            using (var namePaint = new SKPaint
            {
                IsAntialias = true,
                Color = brandNavy,
                TextSize = 80,
                Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold)
            })
            {
                float nameWidth = rightW;
                // Prefer single line, but allow wrap to 2 lines when needed (no forced ellipsis)
                float nameHeight = DrawAdaptiveText(canvas, displayName, rightX, cursorY, nameWidth, namePaint,
                    preferredSize: 80f, minSize: 44f, maxLines: 2, lineSpacing: 8f, preferSingleLine: true, singleLineEllipsize: false);
                cursorY += nameHeight;
            }

            // Add more generous spacing between name and role
            cursorY += 35;

            // Job Title - blue and prominent, also using full width
            using (var rolePaint = new SKPaint
            {
                IsAntialias = true,
                Color = brandBlue,
                TextSize = 52,
                Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold)
            })
            {
                var roleText = string.IsNullOrWhiteSpace(jobTitle) ? "" : jobTitle.Trim();
                if (!string.IsNullOrEmpty(roleText))
                {
                    float roleWidth = rightW;
                    float roleHeight = DrawAdaptiveText(canvas, roleText, rightX, cursorY, roleWidth, rolePaint,
                        preferredSize: 52f, minSize: 36f, maxLines: 2, lineSpacing: 6f, preferSingleLine: false);
                    cursorY += roleHeight;
                }
            }

            // Add more generous spacing between role and department
            cursorY += 30;

            // Department - navy and appropriately sized
            using (var deptPaint = new SKPaint
            {
                IsAntialias = true,
                Color = brandNavy,
                TextSize = 40, // Slightly larger department text
                Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Normal)
            })
            {
                var deptText = string.IsNullOrWhiteSpace(department) ? "" : department.Trim();
                if (!string.IsNullOrEmpty(deptText))
                {
                    // Check if department text would overlap with QR code
                    float estimatedDeptHeight = 60f; // Conservative estimate for department text height
                    bool wouldOverlapQR = (cursorY + estimatedDeptHeight) > qrY;
                    
                    // Use full width if no overlap, otherwise constrain to avoid QR
                    float deptWidth = wouldOverlapQR ? (qrX - rightX - 20f) : rightW;
                    
                    cursorY += DrawAdaptiveText(canvas, deptText, rightX, cursorY, deptWidth, deptPaint,
                        preferredSize: 40f, minSize: 28f, maxLines: 2, lineSpacing: 4f, preferSingleLine: false);
                }
            }

            // Render QR
            canvas.DrawBitmap(qr, qrX, qrY);

            // Email/UPN text is intentionally removed per request

            canvas.Flush();
            return badge;
        }

        private static void DrawPhoto(SKCanvas canvas, SKRect rect, SKBitmap? photo, string displayName)
        {
            using var clipPath = new SKPath();
            clipPath.AddRoundRect(rect, 20, 20);
            canvas.Save();
            canvas.ClipPath(clipPath);

            if (photo != null)
            {
                // Center-crop: preserve aspect ratio, fill rect, crop source center
                float destW = rect.Width;
                float destH = rect.Height;
                float destRatio = destW / destH;
                float srcW = photo.Width;
                float srcH = photo.Height;
                float srcRatio = srcW / srcH;

                SKRect src;
                if (srcRatio > destRatio)
                {
                    // Source is wider than dest: crop left/right
                    float newW = srcH * destRatio;
                    float x = (srcW - newW) / 2f;
                    src = new SKRect(x, 0, x + newW, srcH);
                }
                else
                {
                    // Source is taller than dest: crop top/bottom
                    float newH = srcW / destRatio;
                    float y = (srcH - newH) / 2f;
                    src = new SKRect(0, y, srcW, y + newH);
                }

                // Paint a white base to ensure clean corners inside the clip, then draw the photo
                canvas.DrawRect(rect, new SKPaint { Color = SKColors.White, IsAntialias = true });
                canvas.DrawBitmap(photo, src, rect, new SKPaint { FilterQuality = SKFilterQuality.High, IsAntialias = true });
            }
            else
            {
                // Fallback: light gray background with initials
                // Draw the placeholder inset so it doesn't fill the entire photo panel
                float inset = 24f; // visual padding inside the photo panel
                var phRect = rect;
                phRect.Inflate(-inset, -inset);
                using var bgPaint = new SKPaint { Color = new SKColor(235, 235, 235), IsAntialias = true };
                canvas.DrawRoundRect(phRect, 16f, 16f, bgPaint);

                var initials = GetInitials(displayName);
                // Scale initials relative to the available placeholder height
                float targetH = phRect.Height * 0.35f; // ~35% of height
                float textSize = Math.Max(48f, Math.Min(110f, targetH));
                using var textPaint = new SKPaint
                {
                    IsAntialias = true,
                    Color = new SKColor(120, 120, 120),
                    TextSize = textSize,
                    TextAlign = SKTextAlign.Center,
                    Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold)  // Arial per brand guidelines
                };
                var cx = phRect.MidX;
                var cy = phRect.MidY + textSize * 0.35f; // baseline adjust to center-ish
                canvas.DrawText(initials, cx, cy, textPaint);
            }

            canvas.Restore();

            // Draw border stroke that fully covers the photo edge with no gaps
            var borderColor = new SKColor(0x00, 0x90, 0xD0); // #0090D0
            float stroke = PhotoBorderStroke;
            
            // Inflate the border rect slightly inward so the stroke covers any edge gaps
            var borderRect = rect;
            borderRect.Inflate(-1f, -1f); // 1px inward to ensure stroke covers photo edge completely
            float borderRadius = 19f; // slightly smaller radius to match the inset
            
            using var borderPaint = new SKPaint
            {
                IsAntialias = true,
                Style = SKPaintStyle.Stroke,
                Color = borderColor,
                StrokeWidth = stroke,
                StrokeJoin = SKStrokeJoin.Round,
                StrokeCap = SKStrokeCap.Round
            };
            using var borderPath = new SKPath();
            borderPath.AddRoundRect(borderRect, borderRadius, borderRadius);
            canvas.DrawPath(borderPath, borderPaint);
        }

        private static float DrawAdaptiveText(SKCanvas canvas, string text, float x, float y, float maxWidth, SKPaint paint,
            float preferredSize, float minSize, int maxLines, float lineSpacing, bool preferSingleLine, bool singleLineEllipsize = false)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0f;

            preferredSize = Math.Max(preferredSize, minSize);
            float size = preferredSize;
            paint.TextSize = size;

            if (preferSingleLine)
            {
                float measured = paint.MeasureText(text);
                while (measured > maxWidth && size > minSize)
                {
                    size -= 2f;
                    paint.TextSize = size;
                    measured = paint.MeasureText(text);
                }

                if (measured <= maxWidth)
                {
                    canvas.DrawText(text, x, y + size, paint);
                    return size + 2f;
                }

                if (singleLineEllipsize)
                {
                    paint.TextSize = minSize;
                    string clipped = EllipsizeToWidth(text, paint, maxWidth);
                    canvas.DrawText(clipped, x, y + minSize, paint);
                    return minSize + 2f;
                }

                paint.TextSize = size;
            }
            else
            {
                paint.TextSize = preferredSize;
            }

            // Fallback to wrapping using current size down to minSize
            float startSize = Math.Min(paint.TextSize, preferredSize);
            paint.TextSize = startSize;
            return DrawWrappedText(canvas, text, x, y, maxWidth, paint, maxLines, minSize, lineSpacing);
        }

        private static float DrawWrappedText(SKCanvas canvas, string text, float x, float y, float maxWidth, SKPaint paint, int maxLines = 2, float minSize = 16f, float lineSpacing = 2f)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0f;

            // Try from current size down to minSize until the wrapped lines fit within maxLines
            float size = paint.TextSize;
            while (size >= minSize)
            {
                paint.TextSize = size;
                var lines = WrapWords(text, paint, maxWidth);
                if (lines.Count <= maxLines)
                {
                    // Draw lines
                    float consumed = 0f;
                    for (int i = 0; i < lines.Count; i++)
                    {
                        var line = lines[i];
                        canvas.DrawText(line, x, y + consumed + size, paint);
                        consumed += size + (i < lines.Count - 1 ? lineSpacing : 0f);
                    }
                    return consumed;
                }
                size -= 2f;
            }

            // If still too many lines at min size, render two lines and truncate with ellipsis
            paint.TextSize = minSize;
            var all = WrapWords(text, paint, maxWidth);
            if (all.Count <= maxLines)
            {
                float consumed = 0f;
                for (int i = 0; i < all.Count; i++)
                {
                    var line = all[i];
                    canvas.DrawText(line, x, y + consumed + minSize, paint);
                    consumed += minSize + (i < all.Count - 1 ? lineSpacing : 0f);
                }
                return consumed;
            }

            // Build two lines and ellipsize remainder into second line
            string first = all[0];
            // Combine the rest
            string remaining = string.Join(" ", all.GetRange(1, all.Count - 1));
            string second = EllipsizeToWidth(remaining, paint, maxWidth);

            canvas.DrawText(first, x, y + minSize, paint);
            canvas.DrawText(second, x, y + minSize + lineSpacing + minSize, paint);
            return minSize * 2 + lineSpacing;
        }

        private static System.Collections.Generic.List<string> WrapWords(string text, SKPaint paint, float maxWidth)
        {
            var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var lines = new System.Collections.Generic.List<string>();
            string current = string.Empty;
            foreach (var w in words)
            {
                string candidate = string.IsNullOrEmpty(current) ? w : current + " " + w;
                if (paint.MeasureText(candidate) <= maxWidth)
                {
                    current = candidate;
                    continue;
                }

                // Push current line if any and try to place the long word on the next line
                if (!string.IsNullOrEmpty(current))
                {
                    lines.Add(current);
                    current = string.Empty;
                }

                // If the single word is still wider than maxWidth and contains hyphens, split at a hyphen that fits
                if (paint.MeasureText(w) > maxWidth && w.Contains('-'))
                {
                    string remaining = w;
                    while (!string.IsNullOrEmpty(remaining) && paint.MeasureText(remaining) > maxWidth && remaining.Contains('-'))
                    {
                        int breakPos = -1;
                        float measuredUpTo = 0f;
                        for (int i = 0; i < remaining.Length; i++)
                        {
                            char ch = remaining[i];
                            string prefix = remaining.Substring(0, i + 1);
                            float m = paint.MeasureText(prefix);
                            if (m <= maxWidth)
                            {
                                measuredUpTo = m;
                                if (ch == '-') breakPos = i + 1; // include hyphen on the line end
                            }
                            else break;
                        }
                        if (breakPos > 0)
                        {
                            // Emit prefix (with trailing hyphen)
                            lines.Add(remaining.Substring(0, breakPos));
                            remaining = remaining.Substring(breakPos);
                        }
                        else
                        {
                            break; // cannot split at hyphen within width; let outer shrink logic handle
                        }
                    }
                    // Whatever remains goes to current (even if still long; DrawWrappedText will shrink as needed)
                    current = remaining;
                }
                else
                {
                    current = w; // Let DrawWrappedText reduce size if needed
                }
            }
            if (!string.IsNullOrEmpty(current)) lines.Add(current);
            return lines;
        }

        private static string EllipsizeToWidth(string text, SKPaint paint, float maxWidth)
        {
            const string ellipsis = "…";
            if (paint.MeasureText(text) <= maxWidth) return text;
            string t = text;
            // Remove characters until it fits
            while (t.Length > 0 && paint.MeasureText(t + ellipsis) > maxWidth)
            {
                t = t.Substring(0, t.Length - 1);
            }
            return t.Length == 0 ? ellipsis : t + ellipsis;
        }

        private static float MmToPx(float mm)
        {
            return mm * Dpi / 25.4f;
        }

        private static string GetInitials(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "?";
            var parts = name.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1) return parts[0].Substring(0, Math.Min(2, parts[0].Length)).ToUpperInvariant();
            return ($"{parts[0][0]}{parts[^1][0]}").ToUpperInvariant();
        }

        public static string BuildQr(string upn)
        {
            // Requested format: "{" + "C" + local-part of UPN (before '@') + "}"
            var local = upn;
            var at = upn.IndexOf('@');
            if (at > 0) local = upn.Substring(0, at);
            return "{" + "C" + local + "}";
        }

        public static SKBitmap? LoadSkBitmapFromStream(Stream stream)
        {
            try
            {
                return SKBitmap.Decode(stream);
            }
            catch
            {
                return null;
            }
        }

        public static SKBitmap GenerateQrBitmap(string content, int size)
        {
            using var generator = new QRCodeGenerator();
            using var data = generator.CreateQrCode(content, QRCodeGenerator.ECCLevel.H);
            int modules = data.ModuleMatrix.Count; // number of data modules per side
            // Prior behavior used quiet zones (4 modules each side). To keep the visual size the same,
            // scale the no-quiet bitmap down by modules/(modules + 8)
            int targetSize = (int)Math.Round(size * (double)modules / (modules + 8));
            using var qrCode = new QRCode(data);
            using var bmp = qrCode.GetGraphic(20, System.Drawing.Color.Black, System.Drawing.Color.Transparent, drawQuietZones: false);
            return ComposeQrBitmap(bmp, size, size); // draw to full canvas so bottom aligns exactly
        }

        public static SKBitmap GenerateQrBitmap(string content, int size, SKColor darkColor)
        {
            using var generator = new QRCodeGenerator();
            using var data = generator.CreateQrCode(content, QRCodeGenerator.ECCLevel.H);
            int modules = data.ModuleMatrix.Count;
            int targetSize = (int)Math.Round(size * (double)modules / (modules + 8));
            using var qrCode = new QRCode(data);
            var dark = System.Drawing.Color.FromArgb(darkColor.Alpha, darkColor.Red, darkColor.Green, darkColor.Blue);
            using var bmp = qrCode.GetGraphic(20, dark, System.Drawing.Color.Transparent, drawQuietZones: false);
            return ComposeQrBitmap(bmp, size, size); // draw to full canvas so bottom aligns exactly
        }

        private static SKBitmap ComposeQrBitmap(System.Drawing.Bitmap bmp, int canvasSize, int codeSize)
        {
            using var ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;
            var sk = SKBitmap.Decode(ms);
            if (sk == null)
            {
                return new SKBitmap(canvasSize, canvasSize); // fallback blank
            }
            var resized = new SKBitmap(canvasSize, canvasSize);
            using var canvas = new SKCanvas(resized);
            canvas.Clear(SKColors.Transparent);
            var src = new SKRect(0, 0, sk.Width, sk.Height);
            var dst = new SKRect(0, 0, canvasSize, canvasSize); // fill entire canvas, no transparent padding
            var paint = new SKPaint { FilterQuality = SKFilterQuality.None, IsAntialias = false };
            canvas.DrawBitmap(sk, src, dst, paint);
            canvas.Flush();
            sk.Dispose();
            return resized;
        }

        private static SKBitmap? LoadLogoBitmap()
        {
            try
            {
                // 1) Try WPF Resource (pack URI)
                try
                {
                    var uri = new System.Uri("pack://application:,,,/Assets/Logo.png", System.UriKind.Absolute);
                    var sri = System.Windows.Application.GetResourceStream(uri);
                    if (sri != null)
                    {
                        using var s = sri.Stream;
                        var bmp = SKBitmap.Decode(s);
                        if (bmp != null) return bmp;
                    }
                }
                catch { }
                try
                {
                    var uri2 = new System.Uri("pack://application:,,,/Assets/Logo - White.png", System.UriKind.Absolute);
                    var sri2 = System.Windows.Application.GetResourceStream(uri2);
                    if (sri2 != null)
                    {
                        using var s2 = sri2.Stream;
                        var bmp2 = SKBitmap.Decode(s2);
                        if (bmp2 != null) return bmp2;
                    }
                }
                catch { }

                // 2) Try embedded manifest resource (if present)
                var asm = Assembly.GetExecutingAssembly();
                var names = asm.GetManifestResourceNames();
                foreach (var name in names)
                {
                    if (name.EndsWith(".Assets.Logo.png", StringComparison.OrdinalIgnoreCase) ||
                        name.EndsWith("Assets.Logo.png", StringComparison.OrdinalIgnoreCase))
                    {
                        using var s = asm.GetManifestResourceStream(name);
                        if (s != null)
                        {
                            return SKBitmap.Decode(s);
                        }
                    }
                }

                // 3) Fallback to file on disk (developer convenience)
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var p = Path.Combine(baseDir, "Assets", "Logo.png");
                if (File.Exists(p))
                {
                    using var fs = File.OpenRead(p);
                    return SKBitmap.Decode(fs);
                }
                var p2 = Path.Combine(baseDir, "Assets", "Logo - White.png");
                if (File.Exists(p2))
                {
                    using var fs2 = File.OpenRead(p2);
                    return SKBitmap.Decode(fs2);
                }
            }
            catch
            {
                // ignore - no logo available
            }
            return null;
        }

        public static byte[] EncodeToPng(SKBitmap bmp)
        {
            using var data = bmp.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }
    }
}
