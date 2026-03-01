using Microsoft.Playwright;
using OH_Montgomery.Models;

namespace OH_Montgomery.Utils;

/// <summary>DOM extraction helpers for detail pages and captcha.</summary>
public static class DomHelper
{
    public static async Task<string?> ExtractCaptchaBase64Async(IPage page)
    {
        async Task<string?> TryScreenshot(ILocator locator)
        {
            if (await locator.CountAsync() == 0) return null;
            try
            {
                var bytes = await locator.First.ScreenshotAsync();
                return Convert.ToBase64String(bytes);
            }
            catch { return null; }
        }

        var selectors = new[]
        {
            "img[src*='captcha' i]", "img[src*='Captcha']", "img[src*='CAPTCHA']",
            "img[id*='captcha' i]", "img[id*='Captcha']",
            "img[alt*='captcha' i]", "img[alt*='Captcha']",
            "#captchaImage", "#CaptchaImage", "#imgCaptcha",
            ".captcha img", "[id*='captcha' i] img", "[class*='captcha' i] img",
            "canvas[id*='captcha' i]", "canvas[class*='captcha' i]",
            "img[src*='Verify' i]", "img[src*='Security' i]",
            "img[src*='GetCaptcha']", "img[src*='ValidateImage']", "img[src*='ValidateCode']"
        };
        foreach (var sel in selectors)
        {
            var result = await TryScreenshot(page.Locator(sel).First);
            if (result != null) return result;
        }

        foreach (var sel in new[] { "[id*='captcha' i]", "[class*='captcha' i]", "[id*='Captcha']" })
        {
            var result = await TryScreenshot(page.Locator(sel).First);
            if (result != null) return result;
        }

        foreach (var frame in page.Frames)
        {
            if (string.IsNullOrEmpty(frame.Url) || frame.Url == "about:blank") continue;
            try
            {
                foreach (var sel in new[] { "img[src*='captcha']", "img[id*='captcha']", ".captcha img", "canvas" })
                {
                    var el = frame.Locator(sel).First;
                    if (await el.CountAsync() > 0)
                    {
                        var bytes = await el.First.ScreenshotAsync();
                        return Convert.ToBase64String(bytes);
                    }
                }
            }
            catch { }
        }

        try
        {
            var script = @"() => {
                const imgs = Array.from(document.querySelectorAll('img'));
                for (const img of imgs) {
                    if (!img.offsetParent || img.offsetWidth < 50) continue;
                    const w = img.naturalWidth || img.offsetWidth, h = img.naturalHeight || img.offsetHeight;
                    if (w >= 80 && w <= 350 && h >= 30 && h <= 120) {
                        const src = (img.src || '').toLowerCase();
                        if (src.includes('captcha') || src.includes('verify') || src.includes('validate') || src.includes('security') || img.id?.toLowerCase().includes('captcha'))
                            { const r = img.getBoundingClientRect(); return { x: r.x, y: r.y, width: r.width, height: r.height }; }
                    }
                }
                const input = document.querySelector('input[name*=""captcha""], input[id*=""captcha""], input[name*=""Captcha""]');
                if (input) {
                    const prev = input.previousElementSibling || input.parentElement?.querySelector('img');
                    if (prev && prev.tagName === 'IMG') { const r = prev.getBoundingClientRect(); return { x: r.x, y: r.y, width: r.width, height: r.height }; }
                }
                return null;
            }";
            var rect = await page.EvaluateAsync<CaptchaRect?>(script);
            if (rect != null && rect.width > 0 && rect.height > 0)
            {
                var bytes = await page.ScreenshotAsync(new PageScreenshotOptions { Clip = new() { X = (float)rect.x, Y = (float)rect.y, Width = (float)rect.width, Height = (float)rect.height } });
                return Convert.ToBase64String(bytes);
            }
        }
        catch { }

        return null;
    }

    public static async Task<string> GetDetailValue(IPage page, string label)
    {
        try
        {
            var group = page.Locator("div.input-group").Filter(new() { Has = page.Locator($"span.informationTitle:has-text('{label}')") }).First;
            if (await group.CountAsync() > 0)
            {
                var dataEl = group.Locator(".informationData");
                if (await dataEl.CountAsync() > 0)
                    return (await dataEl.InnerTextAsync()).Trim();
            }
        }
        catch { }
        return "";
    }

    /// <summary>Extract multi-value field (e.g. MORTGAGEES/GRANTEES) from Document panel. Returns values joined with "; ".</summary>
    public static async Task<string> GetDetailValueList(IPage page, string label)
    {
        try
        {
            var dataEl = page.Locator($"div.input-group:has(span.informationTitle:has-text('{label}')) >> .informationData").First;
            if (await dataEl.CountAsync() == 0) return "";

            var parts = new List<string>();
            var pTags = dataEl.Locator("> p");
            var pCount = await pTags.CountAsync();
            if (pCount > 0)
            {
                for (var i = 0; i < pCount; i++)
                {
                    var t = (await pTags.Nth(i).InnerTextAsync()).Trim();
                    if (!string.IsNullOrWhiteSpace(t)) parts.Add(t);
                }
            }
            else
            {
                var childTags = dataEl.Locator("p, div");
                var n = await childTags.CountAsync();
                if (n > 0)
                {
                    for (var i = 0; i < n; i++)
                    {
                        var t = (await childTags.Nth(i).InnerTextAsync()).Trim();
                        if (!string.IsNullOrWhiteSpace(t)) parts.Add(t);
                    }
                }
                else
                {
                    var full = (await dataEl.InnerTextAsync()).Trim();
                    if (!string.IsNullOrWhiteSpace(full))
                    {
                        var lineParts = full.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .Where(s => !string.IsNullOrWhiteSpace(s));
                        parts.AddRange(lineParts);
                    }
                }
            }
            return string.Join("; ", parts);
        }
        catch { }
        return "";
    }

    /// <summary>Get Grantee field by trying document-type-specific labels (GRANTEES, MORTGAGEES, ASSIGNEES, etc.).</summary>
    public static async Task<string> GetGranteeValueList(IPage page)
    {
        var labels = new[] { "GRANTEES", "MORTGAGEES", "ASSIGNEES", "GRANTEE", "MORTGAGEE", "ASSIGNEE" };
        foreach (var label in labels)
        {
            var value = await GetDetailValueList(page, label);
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        return "";
    }

    public static async Task<string> GetRemarksList(IPage page)
    {
        try
        {
            var parts = new List<string>();
            var noteGroups = page.Locator("div.input-group").Filter(new() { Has = page.Locator("span.informationTitle:has-text('NOTE')") });
            for (var i = 0; i < await noteGroups.CountAsync(); i++)
            {
                var dataEl = noteGroups.Nth(i).Locator(".informationData");
                if (await dataEl.CountAsync() > 0)
                {
                    var childTags = dataEl.Locator("p, div");
                    var n = await childTags.CountAsync();
                    if (n > 0)
                    {
                        for (var j = 0; j < n; j++)
                        {
                            var t = (await childTags.Nth(j).InnerTextAsync()).Trim();
                            if (!string.IsNullOrEmpty(t)) parts.Add(t);
                        }
                    }
                    else
                    {
                        var t = (await dataEl.InnerTextAsync()).Trim();
                        if (!string.IsNullOrEmpty(t)) parts.Add(t);
                    }
                }
            }
            return string.Join(";", parts.Where(s => !string.IsNullOrWhiteSpace(s)));
        }
        catch { }
        return "";
    }
}
