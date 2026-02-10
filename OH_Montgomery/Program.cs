using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Playwright;
using OH_Montgomery;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Tiff;

class Program
{
    static void LogInput(ActorInput input)
    {
        Console.WriteLine("[Input] Parsed config:");
        Console.WriteLine($"  searchMode={input.SearchMode}, exportMode={input.ExportMode}, sortOrder={input.SortOrder}");
        Console.WriteLine($"  startDate={input.StartDate ?? "(null)"}, endDate={input.EndDate ?? "(null)"}");
        Console.WriteLine($"  side={input.Side ?? "(null)"}, indexTypes=[{string.Join(", ", input.IndexTypes ?? Array.Empty<string>())}]");
        Console.WriteLine($"  twoCaptchaApiKey={(string.IsNullOrEmpty(input.TwoCaptchaApiKey) ? "(not set)" : "[set]")}");
    }

    [STAThread]
    static void Main(string[] args)
    {
        RunAsync().GetAwaiter().GetResult();
    }

    static async Task RunAsync()
    {
        Console.WriteLine("=== OH_Montgomery RISS Scraper (Apify Input) ===");
        Console.WriteLine();

        var input = ApifyHelper.GetInput<ActorInput>();
        if (input == null)
        {
            Console.WriteLine("ERROR: Failed to load input.");
            Environment.Exit(1);
        }

        input.SearchMode = input.SearchMode?.Trim() ?? "ByDate";
        input.ExportMode = input.ExportMode?.Trim() ?? "ExportDataOnly";
        input.SortOrder = input.SortOrder?.Trim() ?? "FILE DATE ASC";
        input.Side = string.IsNullOrWhiteSpace(input.Side) ? "Both" : input.Side.Trim();
        input.IndexTypes ??= ["Deeds", "Mortgages"];
        if (input.IndexTypes.Length == 0) input.IndexTypes = ["Deeds", "Mortgages"];

        LogInput(input);

        try
        {
            input.ValidateByName();
            input.ValidateByType();
            input.ValidateByMunicipality();
            input.ValidateBySTR();
            input.ValidateBySubdivision();
            input.ValidateByInstrument();
            input.ValidateByBookPage();
            input.ValidateByFiche();
            input.ValidateByPre1980();
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Environment.Exit(1);
        }

        var needsCaptcha = string.Equals(input.SearchMode, "ByDate", StringComparison.OrdinalIgnoreCase)
            || string.Equals(input.SearchMode, "ByInstrument", StringComparison.OrdinalIgnoreCase);
        if (needsCaptcha && string.IsNullOrWhiteSpace(input.TwoCaptchaApiKey) && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("TWO_CAPTCHA_API_KEY")))
        {
            Console.WriteLine("ERROR: By Date and By Instrument modes require captcha. Provide 'twoCaptchaApiKey' in input or set TWO_CAPTCHA_API_KEY env var.");
            Environment.Exit(1);
        }

        Console.WriteLine($"[1] Config loaded: SearchMode={input.SearchMode}, ExportMode={input.ExportMode}");
        Console.WriteLine();

        try
        {
            // --- Phase 2: Launch Playwright and navigate ---
            Console.WriteLine("[2] Checking Playwright...");
            Microsoft.Playwright.Program.Main(["install", "chromium"]);
            Console.WriteLine("    Chromium ready.");

            Console.WriteLine("[3] Launching Chromium...");
            var playwright = await Playwright.CreateAsync();
            var isApify = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("APIFY_CONTAINER_PORT"));
            var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = isApify,
                Timeout = 60000,
                Args = new[]
                {
                    "--disable-popup-blocking",
                    "--disable-gpu",
                    "--disable-dev-shm-usage",
                    "--disable-extensions",
                    "--disable-background-networking",
                    "--disable-default-apps",
                    "--disable-sync",
                    "--no-first-run",
                    "--disable-software-rasterizer",
                    "--disable-features=VizDisplayCompositor",
                    "--disk-cache-size=0",
                    "--media-cache-size=0",
                    "--mute-audio"
                }
            });

            Console.WriteLine("[4-11] Creating context and loading results...");
            IBrowserContext context; IPage page; int rowCount; int fileNumberColIndex; int fileDateColIndex; int imageColIndex; string tableSelector; bool isFormSubmitMode;
            try
            {
                (context, page, rowCount, fileNumberColIndex, fileDateColIndex, imageColIndex, tableSelector, isFormSubmitMode) = await CreateContextAndLoadResultsAsync(browser, input, needsCaptcha);
            }
            catch (Exception ex)
            {
                var is500Limit = ex is InvalidOperationException && ex.Message.Contains("Record Count Exceeds", StringComparison.OrdinalIgnoreCase);
                if (!is500Limit) Console.WriteLine($"    Error: {ex.Message}");
                else Console.WriteLine("    Exiting.");
                await browser.CloseAsync();
                return;
            }
            Console.WriteLine("    Grid loaded.");
            Console.WriteLine();

            // --- Scrape ---
            var grid = page.Locator(tableSelector);
            if (await grid.CountAsync() > 0)
            {
                var dataRows = grid.Locator("tbody tr");
                Console.WriteLine($"    Found {rowCount} rows.");

                if (rowCount == 0)
                {
                    Console.WriteLine("    No records. Exiting.");
                    await browser.CloseAsync();
                    return;
                }

                var dateForFilename = input.GetDateForFilename();
                var baseDir = Directory.GetCurrentDirectory();
                if (baseDir.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar))
                {
                    var parts = baseDir.Split(Path.DirectorySeparatorChar);
                    var binIdx = Array.FindLastIndex(parts, p => string.Equals(p, "bin", StringComparison.OrdinalIgnoreCase));
                    if (binIdx > 0) baseDir = string.Join(Path.DirectorySeparatorChar, parts.Take(binIdx));
                }
                var shouldDownloadImages = input.ShouldExportImages && imageColIndex >= 0;
                if (shouldDownloadImages)
                    try { await context.GrantPermissionsAsync(new[] { "clipboard-read", "clipboard-write", "storage-access" }, new BrowserContextGrantPermissionsOptions { Origin = "https://onbase.mcohio.org" }); } catch { }

                Console.WriteLine($"    ExportMode={input.ExportMode}, imageColIndex={imageColIndex}, shouldDownloadImages={shouldDownloadImages}");
                Console.WriteLine($"    Processing rows: image + detail per row, batch 10 to CSV (Images/{dateForFilename}/...)");
                var buffer = new List<Record>();
                var kvStore = Path.Combine(baseDir, "apify_storage", "key_value_store");
                Directory.CreateDirectory(kvStore);
                var outputPath = Path.Combine(kvStore, $"OH_Montgomery_{dateForFilename}.csv");
                var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = "|", HasHeaderRecord = true, ShouldQuote = args => true };
                await using (var writer = new StreamWriter(outputPath, append: false))
                using (var csv = new CsvWriter(writer, csvConfig))
                {
                    for (var r = 0; r < rowCount; r++)
                    {
                        if ((r + 1) % 10 == 0) Console.WriteLine($"    Row {r + 1}/{rowCount}...");
                        grid = page.Locator(tableSelector);
                        dataRows = grid.Locator("tbody tr");
                        var row = dataRows.Nth(r);
                        var cells = row.Locator("td");
                        var cellCount = await cells.CountAsync();
                        if (cellCount == 0) continue;

                        var record = new Record();
                        record.RecordingDate = fileDateColIndex >= 0 && fileDateColIndex < cellCount ? (await cells.Nth(fileDateColIndex).InnerTextAsync()).Trim() : "";

                        // By Book/Page: table has ROW, TYPE, BOOK or BOOK-PREFIX, PAGE, IMAGE(s); no detail page
                        if (input.IsByBookPage && cellCount >= 4)
                        {
                            record.DocumentType = (await cells.Nth(1).InnerTextAsync()).Trim();
                            record.Book = (await cells.Nth(2).InnerTextAsync()).Trim();
                            record.Page = (await cells.Nth(3).InnerTextAsync()).Trim();
                            record.BookType = record.DocumentType;
                            record.DocumentNumber = $"{record.Book}-{record.Page}".TrimEnd('-');
                        }
                        // By Pre-1980: table has ROW, TYPE, INSTRUMENT, IMAGE(s); no detail page
                        else if (input.IsByPre1980 && cellCount >= 4)
                        {
                            record.DocumentType = (await cells.Nth(1).InnerTextAsync()).Trim();
                            record.DocumentNumber = (await cells.Nth(2).InnerTextAsync()).Trim();
                            record.BookType = record.DocumentType;
                        }

                        var imageLinks = new List<string>();

                        if (shouldDownloadImages)
                        {
                            var imageCell = cells.Nth(imageColIndex);
                            var cameraLink = imageCell.Locator("a").First;
                            if (await cameraLink.CountAsync() > 0)
                            {
                                IPage? imgPage = null;
                                try
                                {
                                    await page.BringToFrontAsync();
                                    await page.WaitForTimeoutAsync(300);
                                    imgPage = await context.RunAndWaitForPageAsync(async () => { await cameraLink.ClickAsync(); }, new BrowserContextRunAndWaitForPageOptions { Timeout = 30000 });
                                    await imgPage.BringToFrontAsync();
                                    imgPage.SetDefaultTimeout(35000);
                                    await imgPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                                    await imgPage.WaitForTimeoutAsync(2000);
                                    await imgPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                                    await imgPage.WaitForTimeoutAsync(6000);

                                    // Some image pages first display a "Document Search Results" list in an iframe.
                                    // In that case we must double-click each row (folder) in sequence to open the actual image viewer.
                                    try
                                    {
                                        // Find iframe that contains the document select list
                                        IFrame? selectListFrame = null;
                                        foreach (var fr in imgPage.Frames)
                                        {
                                            try
                                            {
                                                var hasSelectList = await fr.EvaluateAsync<bool>(@"() => !!document.querySelector('#DocumentSelectList, #primaryHitlist_grid')");
                                                if (hasSelectList)
                                                {
                                                    selectListFrame = fr;
                                                    break;
                                                }
                                            }
                                            catch { }
                                        }

                                        if (selectListFrame != null)
                                        {
                                            // Get all rows in the result grid, each row represents one "folder"
                                            var folderRows = selectListFrame.Locator("#primaryHitlist_grid tbody tr");
                                            var folderCount = await folderRows.CountAsync();

                                            for (var frIndex = 0; frIndex < folderCount; frIndex++)
                                            {
                                                var rowFolder = folderRows.Nth(frIndex);
                                                var firstCell = rowFolder.Locator("td").First;

                                                if (await firstCell.CountAsync() == 0)
                                                    continue;

                                                // Double-click to open the folder / document
                                                try
                                                {
                                                    await firstCell.DblClickAsync();
                                                }
                                                catch
                                                {
                                                    // If DblClick fails, fall back to two single clicks
                                                    try
                                                    {
                                                        await firstCell.ClickAsync();
                                                        await imgPage.WaitForTimeoutAsync(200);
                                                        await firstCell.ClickAsync();
                                                    }
                                                    catch { }
                                                }

                                                // Wait for image viewer to appear
                                                IFrame? viewerFrameForFolder = null;
                                                for (var findAttempt = 0; findAttempt < 20; findAttempt++)
                                                {
                                                    foreach (var fr in imgPage.Frames)
                                                    {
                                                        try
                                                        {
                                                            var hasViewer = await fr.EvaluateAsync<bool>(@"() => !!document.querySelector('#htmlViewer, img.document-image')");
                                                            if (hasViewer) { viewerFrameForFolder = fr; break; }
                                                        }
                                                        catch { }
                                                    }
                                                    if (viewerFrameForFolder != null) break;
                                                    await imgPage.WaitForTimeoutAsync(1000);
                                                }

                                                // If we have an image viewer after the double-click,
                                                // process all pages for this folder before moving to the next one.
                                                if (viewerFrameForFolder == null)
                                                    continue;

                                                var frameForFolder = viewerFrameForFolder;
                                                await frameForFolder.WaitForSelectorAsync("img.document-image", new FrameWaitForSelectorOptions { State = WaitForSelectorState.Visible, Timeout = 25000 });
                                                for (var waitImg = 0; waitImg < 25; waitImg++)
                                                {
                                                    var ready = await frameForFolder.EvaluateAsync<bool>(@"() => {
                                                        const img = document.querySelector('img.document-image');
                                                        return !!(img && img.complete && img.naturalWidth > 0);
                                                    }");
                                                    if (ready) break;
                                                    await imgPage.WaitForTimeoutAsync(500);
                                                }
                                                await imgPage.WaitForTimeoutAsync(2000);

                                                // Base name for the current folder to avoid overwriting files between folders
                                                var baseNameForFolder = $"row{r + 1}_sel{frIndex + 1}";

                                                var canvasSingleScriptFolder = @"() => {
                                                    const img = document.querySelector('img.document-image');
                                                    if (!img?.src || img.naturalWidth === 0) return null;
                                                    const canvas = document.createElement('canvas');
                                                    canvas.width = img.naturalWidth;
                                                    canvas.height = img.naturalHeight;
                                                    const ctx = canvas.getContext('2d');
                                                    if (!ctx) return null;
                                                    ctx.drawImage(img, 0, 0);
                                                    const dataUrl = canvas.toDataURL('image/png');
                                                    return dataUrl ? dataUrl.split(',')[1] : null;
                                                }";

                                                var base64ListFolder = new List<string>();
                                                var b64Folder = await frameForFolder.EvaluateAsync<string?>(canvasSingleScriptFolder);
                                                if (!string.IsNullOrEmpty(b64Folder)) base64ListFolder.Add(b64Folder);

                                                var nextBtnFolder = frameForFolder.Locator("#next-page .button-item");
                                                while (true)
                                                {
                                                    var disabled = await nextBtnFolder.GetAttributeAsync("aria-disabled");
                                                    if (string.Equals(disabled, "true", StringComparison.OrdinalIgnoreCase)) break;
                                                    await nextBtnFolder.ClickAsync();
                                                    await imgPage.WaitForTimeoutAsync(2000);
                                                    for (var w = 0; w < 15; w++)
                                                    {
                                                        var imgReady = await frameForFolder.EvaluateAsync<bool>(@"() => {
                                                            const img = document.querySelector('img.document-image');
                                                            return !!(img && img.complete && img.naturalWidth > 0);
                                                        }");
                                                        if (imgReady) break;
                                                        await imgPage.WaitForTimeoutAsync(400);
                                                    }
                                                    b64Folder = await frameForFolder.EvaluateAsync<string?>(canvasSingleScriptFolder);
                                                    if (!string.IsNullOrEmpty(b64Folder)) base64ListFolder.Add(b64Folder);
                                                }

                                                var base64ArrFolder = base64ListFolder.ToArray();
                                                if (base64ArrFolder.Length > 0)
                                                {
                                                    for (var i = 0; i < base64ArrFolder.Length; i++)
                                                    {
                                                        if (string.IsNullOrEmpty(base64ArrFolder[i])) continue;
                                                        try
                                                        {
                                                            var pngBytes = Convert.FromBase64String(base64ArrFolder[i]);
                                                            var bytes = ConvertPngToTiff(pngBytes);
                                                            var fileName = base64ArrFolder.Length > 1 ? $"{baseNameForFolder}-{i + 1}.tif" : $"{baseNameForFolder}.tif";
                                                            var kvKey = $"Images/{dateForFilename}/{baseNameForFolder}/{fileName}";
                                                            await ApifyHelper.SaveImageAsync(kvKey, bytes);
                                                            if (input.ExportAll) imageLinks.Add(ApifyHelper.GetRecordUrl(kvKey));
                                                            Console.WriteLine($"      Saved: {kvKey}");
                                                        }
                                                        catch (Exception ex) { Console.WriteLine($"      Save error: {ex.Message}"); }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }

                                    IFrame? viewerFrame = null;
                                    for (var findAttempt = 0; findAttempt < 20; findAttempt++)
                                    {
                                        foreach (var fr in imgPage.Frames)
                                        {
                                            try
                                            {
                                                var hasViewer = await fr.EvaluateAsync<bool>(@"() => !!document.querySelector('#htmlViewer, img.document-image')");
                                                if (hasViewer) { viewerFrame = fr; break; }
                                            }
                                            catch { }
                                        }
                                        if (viewerFrame != null) break;
                                        await imgPage.WaitForTimeoutAsync(1000);
                                    }
                                    var frame = viewerFrame ?? imgPage.MainFrame;
                                    await frame.WaitForSelectorAsync("img.document-image", new FrameWaitForSelectorOptions { State = WaitForSelectorState.Visible, Timeout = 25000 });
                                    for (var waitImg = 0; waitImg < 25; waitImg++)
                                    {
                                        var ready = await frame.EvaluateAsync<bool>(@"() => {
                                            const img = document.querySelector('img.document-image');
                                            return !!(img && img.complete && img.naturalWidth > 0);
                                        }");
                                        if (ready) break;
                                        await imgPage.WaitForTimeoutAsync(500);
                                    }
                                    await imgPage.WaitForTimeoutAsync(2000);
                                    var pageTitle = await imgPage.TitleAsync();
                                    var fileNumberText = fileNumberColIndex >= 0 && fileNumberColIndex < cellCount ? (await cells.Nth(fileNumberColIndex).InnerTextAsync()).Trim() : "";
                                    string baseName;
                                    if (input.IsByBookPage && !string.IsNullOrEmpty(record.Book))
                                        baseName = Regex.Replace($"{record.Book}_{record.Page}".TrimEnd('_'), @"[^\w\-]", "_");
                                    else if (input.IsByPre1980 && !string.IsNullOrEmpty(record.DocumentNumber))
                                        baseName = Regex.Replace(record.DocumentNumber.Trim(), @"[^\w\-]", "_");
                                    else
                                    {
                                        var titleMatch = Regex.Match(pageTitle ?? "", @"Instrument Number\s+([\w\-]+)|[\w\-]*?(\d{5,})$");
                                        baseName = titleMatch.Success ? (titleMatch.Groups[1].Value ?? titleMatch.Groups[2].Value ?? "") : "";
                                        if (string.IsNullOrWhiteSpace(baseName)) baseName = fileNumberText;
                                        baseName = Regex.Replace(baseName?.Trim() ?? fileNumberText, @"[^\w\-]", "_");
                                        if (string.IsNullOrWhiteSpace(baseName)) baseName = $"row{r + 1}";
                                    }
                                    var canvasSingleScript = @"() => {
                                        const img = document.querySelector('img.document-image');
                                        if (!img?.src || img.naturalWidth === 0) return null;
                                        const canvas = document.createElement('canvas');
                                        canvas.width = img.naturalWidth;
                                        canvas.height = img.naturalHeight;
                                        const ctx = canvas.getContext('2d');
                                        if (!ctx) return null;
                                        ctx.drawImage(img, 0, 0);
                                        const dataUrl = canvas.toDataURL('image/png');
                                        return dataUrl ? dataUrl.split(',')[1] : null;
                                    }";
                                    var base64List = new List<string>();
                                    var b64 = await frame.EvaluateAsync<string?>(canvasSingleScript);
                                    if (!string.IsNullOrEmpty(b64)) base64List.Add(b64);
                                    var nextBtn = frame.Locator("#next-page .button-item");
                                    while (true)
                                    {
                                        var disabled = await nextBtn.GetAttributeAsync("aria-disabled");
                                        if (string.Equals(disabled, "true", StringComparison.OrdinalIgnoreCase)) break;
                                        await nextBtn.ClickAsync();
                                        await imgPage.WaitForTimeoutAsync(2000);
                                        for (var w = 0; w < 15; w++)
                                        {
                                            var imgReady = await frame.EvaluateAsync<bool>(@"() => {
                                                const img = document.querySelector('img.document-image');
                                                return !!(img && img.complete && img.naturalWidth > 0);
                                            }");
                                            if (imgReady) break;
                                            await imgPage.WaitForTimeoutAsync(400);
                                        }
                                        b64 = await frame.EvaluateAsync<string?>(canvasSingleScript);
                                        if (!string.IsNullOrEmpty(b64)) base64List.Add(b64);
                                    }
                                    var base64Arr = base64List.ToArray();
                                    if (base64Arr.Length > 0)
                                    {
                                        for (var i = 0; i < base64Arr.Length; i++)
                                        {
                                            if (string.IsNullOrEmpty(base64Arr[i])) continue;
                                            try
                                            {
                                                var pngBytes = Convert.FromBase64String(base64Arr[i]);
                                                var bytes = ConvertPngToTiff(pngBytes);
                                                var fileName = base64Arr.Length > 1 ? $"{baseName}-{i + 1}.tif" : $"{baseName}.tif";
                                                var kvKey = $"Images/{dateForFilename}/{baseName}/{fileName}";
                                                await ApifyHelper.SaveImageAsync(kvKey, bytes);
                                                if (input.ExportAll) imageLinks.Add(ApifyHelper.GetRecordUrl(kvKey));
                                                Console.WriteLine($"      Saved: {kvKey}");
                                            }
                                            catch (Exception ex) { Console.WriteLine($"      Save error: {ex.Message}"); }
                                        }
                                    }
                                }
                                catch (Exception ex) { Console.WriteLine($"    Row {r + 1} image error: {ex.Message}"); }
                                finally
                                {
                                    if (imgPage != null)
                                    {
                                        try { await imgPage.CloseAsync(); } catch { }
                                        await page.BringToFrontAsync();
                                        await page.WaitForTimeoutAsync(500);
                                    }
                                }
                            }
                        }
                        if (input.ExportAll && imageLinks.Count > 0)
                            record.ImageLinks = string.Join("\n", imageLinks.Where(u => !string.IsNullOrWhiteSpace(u)));

                        if (fileNumberColIndex >= 0 && fileNumberColIndex < cellCount)
                        {
                            var fileNumberCell = cells.Nth(fileNumberColIndex);
                            var clickTarget = fileNumberCell.Locator("form input[type=submit], form button, form a, a").First;
                            if (await clickTarget.CountAsync() > 0)
                            {
                                try
                                {
                                    if (isFormSubmitMode)
                                    {
                                        // By Name: form submit navigates in same page; wait for docinfo then GoBack
                                        var navTask = page.WaitForURLAsync(url => url.Contains("docinfo", StringComparison.OrdinalIgnoreCase), new PageWaitForURLOptions { Timeout = 30000 });
                                        await clickTarget.ClickAsync();
                                        await navTask;
                                        await page.WaitForSelectorAsync("span.informationTitle", new PageWaitForSelectorOptions { Timeout = 15000 });
                                        record.DocumentNumber = await GetDetailValue(page, "INSTRUMENT");
                                        record.BookType = await GetDetailValue(page, "INDEX");
                                        record.DocumentType = await GetDetailValue(page, "TYPE");
                                        record.Amount = await GetDetailValue(page, "AMOUNT");
                                        var grantor = await GetDetailValueList(page, "GRANTORS");
                                        if (string.IsNullOrEmpty(grantor)) grantor = await GetDetailValueList(page, "MORTGAGORS");
                                        record.Grantor = grantor;
                                        var grantee = await GetDetailValueList(page, "GRANTEES");
                                        if (string.IsNullOrEmpty(grantee)) grantee = await GetDetailValueList(page, "MORTGAGEES");
                                        record.Reference = await GetDetailValueList(page, "REFERENCES");
                                        record.Remarks = await GetRemarksList(page);
                                        record.Legal = await GetDetailValueList(page, "LEGAL");
                                        record.Property = await GetDetailValueList(page, "PROPERTY");
                                        await page.GoBackAsync(new PageGoBackOptions { WaitUntil = WaitUntilState.NetworkIdle });
                                        await page.WaitForSelectorAsync(tableSelector, new PageWaitForSelectorOptions { Timeout = 15000 });
                                        await page.WaitForTimeoutAsync(500);
                                    }
                                    else
                                    {
                                        // By Date: link opens in new tab; Ctrl+Click
                                        var detailPage = await context.RunAndWaitForPageAsync(async () =>
                                        {
                                            await clickTarget.ClickAsync(new LocatorClickOptions { Modifiers = new[] { KeyboardModifier.Control } });
                                        }, new BrowserContextRunAndWaitForPageOptions { Timeout = 30000 });
                                        try
                                        {
                                            await detailPage.WaitForSelectorAsync("span.informationTitle", new PageWaitForSelectorOptions { Timeout = 15000 });
                                            record.DocumentNumber = await GetDetailValue(detailPage, "INSTRUMENT");
                                            record.BookType = await GetDetailValue(detailPage, "INDEX");
                                            record.DocumentType = await GetDetailValue(detailPage, "TYPE");
                                            record.Amount = await GetDetailValue(detailPage, "AMOUNT");
                                            var grantor = await GetDetailValueList(detailPage, "GRANTORS");
                                            if (string.IsNullOrEmpty(grantor)) grantor = await GetDetailValueList(detailPage, "MORTGAGORS");
                                            record.Grantor = grantor;
                                            var grantee = await GetDetailValueList(detailPage, "GRANTEES");
                                            if (string.IsNullOrEmpty(grantee)) grantee = await GetDetailValueList(detailPage, "MORTGAGEES");
                                            record.Reference = await GetDetailValueList(detailPage, "REFERENCES");
                                            record.Remarks = await GetRemarksList(detailPage);
                                            record.Legal = await GetDetailValueList(detailPage, "LEGAL");
                                            record.Property = await GetDetailValueList(detailPage, "PROPERTY");
                                        }
                                        finally { await detailPage.CloseAsync(); }
                                    }
                                }
                                catch { }
                            }
                        }
                        buffer.Add(record);
                        if ((r + 1) % 10 == 0) { GC.Collect(); GC.WaitForPendingFinalizers(); }

                        if (buffer.Count >= 10)
                        {
                            await ApifyHelper.PushDataAsync<Record>(buffer);
                            await csv.WriteRecordsAsync(buffer);
                            await writer.FlushAsync();
                            Console.WriteLine($"    Batch of {buffer.Count} → CSV + Apify dataset.");
                            buffer.Clear();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }

                    if (buffer.Count > 0)
                    {
                        await ApifyHelper.PushDataAsync<Record>(buffer);
                        await csv.WriteRecordsAsync(buffer);
                        await writer.FlushAsync();
                        Console.WriteLine($"    Final batch of {buffer.Count} → CSV + Apify dataset.");
                    }
                }

                Console.WriteLine($"    Exported {rowCount} records to CSV and Apify dataset.");
                Console.WriteLine($"    Also saved CSV: {outputPath}");
            }
            else
            {
                Console.WriteLine($"    Table ({tableSelector}) not found or empty. Exiting.");
                await browser.CloseAsync();
                return;
            }

            Console.WriteLine();
            Console.WriteLine("[12] Done. Closing browser...");
            await browser.CloseAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex}");
            Environment.Exit(1);
        }
    }

    /// <summary>Convert PNG bytes to TIFF bytes for storage.</summary>
    static byte[] ConvertPngToTiff(byte[] pngBytes)
    {
        using var input = new MemoryStream(pngBytes);
        using var image = Image.Load(input);
        using var output = new MemoryStream();
        image.Save(output, new TiffEncoder());
        return output.ToArray();
    }

    static async Task<string?> ExtractCaptchaBase64Async(IPage page)
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

    static async Task FillCaptchaAsync(IPage page, string captchaCode)
    {
        if (string.IsNullOrWhiteSpace(captchaCode)) return;
        var cap = captchaCode.Trim().ToUpperInvariant();
        var labels = new[] { "CAPTCHA", "Captcha", "Verification Code", "Security Code", "Enter the characters", "Image text" };
        foreach (var label in labels)
        {
            var el = page.GetByLabel(label).Or(page.Locator($"input[name*='captcha' i], input[id*='captcha' i], input[name*='verify' i]").First);
            if (await el.CountAsync() > 0)
            {
                await el.First.FillAsync(cap);
                return;
            }
        }
        var anyCaptchaInput = page.Locator("input[type='text'][name*='captcha' i], input[type='text'][id*='captcha' i]").First;
        if (await anyCaptchaInput.CountAsync() > 0)
            await anyCaptchaInput.First.FillAsync(cap);
    }

    static async Task ClickSidebarTabAsync(IPage page, string searchMode)
    {
        var tabText = searchMode switch
        {
            "ByDate" => "BY DATE",
            "ByName" => "BY NAME",
            "ByType" => "BY TYPE",
            "ByMunicipality" => "BY MUNICIPALITY",
            "BySubdivision" => "BY SUBDIVISION",
            "BySTR" => "BY STR",
            "ByInstrument" => "BY INSTRUMENT",
            "ByBookPage" => "BY BOOK",
            "ByFiche" => "BY FICHE",
            "ByPre1980" => "BY PRE",
            _ => searchMode.ToUpperInvariant()
        };
        // Match tab text (Exact=false: RISS may use "By Date" or "BY DATE"); wait for sidebar
        var tab = page.GetByText(tabText, new() { Exact = false });
        await tab.First.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        await tab.First.ScrollIntoViewIfNeededAsync();
        await tab.First.ClickAsync();
        await page.WaitForTimeoutAsync(500);
    }

    static Task FillByDateForm(IPage page, ActorInput input) => Task.CompletedTask;

    /// <summary>RISS By Name: Side values %=Both, 1=One, 2=Two. Web UI uses Both/Grantor/Grantee.</summary>
    static readonly Dictionary<string, string> SideToRissByName = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Both"] = "%", ["Grantor"] = "1", ["Grantee"] = "2", ["One"] = "1", ["Two"] = "2"
    };

    /// <summary>By Name form uses different Sort values than By Date.</summary>
    static readonly Dictionary<string, string> SortOrderToRissByName = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "CAST([FileDate] as DATE) ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "CAST([FileDate] as DATE) DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    /// <summary>By Name form: LastName*, FirstName, Side, IndexType*, StartDate, EndDate, Sort. No captcha.</summary>
    static async Task FillByNameFormBatchAsync(IPage page, ActorInput input)
    {
        var lastName = input.LastName?.Trim().ToUpperInvariant() ?? "";
        var firstName = input.FirstName?.Trim().ToUpperInvariant() ?? "";
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t) && IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByName.TryGetValue(sortVal, out var sortOpt) ? sortOpt : "CAST([FileDate] as DATE) ASC";
        var sideVal = !string.IsNullOrWhiteSpace(input.Side) && SideToRissByName.TryGetValue(input.Side.Trim(), out var s) ? s : "%";
        var lastNameSearch = NormalizeNameSearch(input.LastNameSearch);
        var firstNameSearch = NormalizeNameSearch(input.FirstNameSearch);

        Console.WriteLine($"[FillByName] LastName={lastName}, FirstName={firstName}, LastNameSearch={lastNameSearch}, FirstNameSearch={firstNameSearch}, Side={sideVal}, IndexTypes=[{string.Join(",", rissIndex)}], StartDate={dateVal}, EndDate={endVal}, Sort={rissSort}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [lastName, firstName, dateVal, endVal, sideVal, idxArr, sortVal, lastNameSearch, firstNameSearch] = args;
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const result = {};
            const lastEl = document.querySelector('#LastName, input[name=""LastName""]');
            const firstEl = document.querySelector('#FirstName, input[name=""FirstName""]');
            const startEl = document.querySelector('#StartDate, input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDate, input[name=""EndDate""]');
            const sideEl = document.querySelector('#Side, select[name=""Side""]');
            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            const lastSearchEl = document.querySelector('#LastName_Search, select[name=""LastName_Search""]');
            const firstSearchEl = document.querySelector('#FirstName_Search, select[name=""FirstName_Search""]');
            if (lastEl) { lastEl.value = lastName || ''; fireChange(lastEl); result.lastName = lastEl.value; }
            if (firstEl) { firstEl.value = firstName || ''; fireChange(firstEl); }
            if (lastSearchEl && lastNameSearch) { lastSearchEl.value = lastNameSearch; fireChange(lastSearchEl); chosenUpdate(lastSearchEl); }
            if (firstSearchEl && firstNameSearch) { firstSearchEl.value = firstNameSearch; fireChange(firstSearchEl); chosenUpdate(firstSearchEl); }
            if (startEl) { startEl.value = dateVal || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = endVal || ''; fireChange(endEl); result.endValue = endEl.value; }
            if (sideEl && sideVal) { sideEl.value = sideVal; fireChange(sideEl); chosenUpdate(sideEl); result.sideValue = sideEl.value; }
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }
            return result;
        }", new object[] { lastName, firstName, dateVal, endVal, sideVal, rissIndex, rissSort, lastNameSearch, firstNameSearch });
        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByName] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByName] Could not log result: {ex.Message}"); }
    }

    static string NormalizeNameSearch(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return "BeginWith";
        var t = v.Trim();
        if (t.Equals("Contains", StringComparison.OrdinalIgnoreCase)) return "Contains";
        if (t.Equals("Exactly", StringComparison.OrdinalIgnoreCase)) return "Exactly";
        if (t.Contains("Begin", StringComparison.OrdinalIgnoreCase)) return "BeginWith";
        return "BeginWith";
    }

    /// <summary>RISS Document Type values (option value -> display text). Use documentTypes with exact text e.g. ["DEED","MORTGAGE","RELEASE OF MORTGAGE"].</summary>
    static readonly Dictionary<string, string> DocumentTypeOptions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["5"] = "AFFIDAVIT (DEED)", ["6"] = "AFFIDAVIT (MORTGAGE)", ["4"] = "AFFIDAVIT (UCC DIVISION)",
        ["15"] = "AGREEMENT", ["13"] = "AGRICULTURAL LIEN", ["9"] = "AMENDMENT FINANCING STATEMENT NON STAND",
        ["10"] = "AMENDMENT FINANCING STATEMENT STANDARD", ["11"] = "AMENDMENT FIXTURE FILING NON STANDARD", ["12"] = "AMENDMENT FIXTURE FILING STANDARD",
        ["17"] = "ANNEXATION", ["18"] = "ARTICLES OF INCORPORATION", ["19"] = "ASSIGNMENT FINANCING STATEMENT NON STAND",
        ["20"] = "ASSIGNMENT FINANCING STATEMENT STANDARD", ["7"] = "ASSIGNMENT FIXTURE FILING NON STANDARD", ["8"] = "ASSIGNMENT FIXTURE FILING STANDARD",
        ["16"] = "ASSIGNMENT OF LAND CONTRACT", ["1"] = "ASSIGNMENT OF LEASE", ["2"] = "ASSIGNMENT OF MORTGAGE",
        ["141"] = "ASSIGNMENT OF TAX LIEN CERTIFICATE", ["3"] = "ASSIGNMENTS OF RENTS", ["22"] = "BANKRUPTCY (DEED)",
        ["23"] = "BANKRUPTCY (MORTGAGE)", ["24"] = "BOND LIEN", ["27"] = "CERTIFICATE OF TRANSFER", ["36"] = "CHILD SUPPORT LIEN",
        ["33"] = "CONDOMINIUM", ["28"] = "CONTINUATION FINANCING STATEMENT NON STD", ["29"] = "CONTINUATION FINANCING STATEMENT STAND",
        ["30"] = "CONTINUATION FIXTURE FILING NON STANDARD", ["31"] = "CONTINUATION FIXTURE FILING STANDARD",
        ["38"] = "CONTINUATION OF UNEMPLOYMENT LIEN", ["25"] = "CONTINUATION OF WORKER'S COMP LIEN", ["37"] = "COURT ENTRY",
        ["40"] = "DECLARATION OF TRUST", ["41"] = "DEED", ["39"] = "DEED CORRECTION", ["42"] = "DISCLAIMER", ["43"] = "DRAWER INDEX",
        ["44"] = "EASEMENT", ["45"] = "ENVIRONMENTAL PROTECTION AGENCY LIEN", ["46"] = "FEDERAL LIEN", ["52"] = "FEDERAL TAX LIEN",
        ["49"] = "FINANCIAL RESPONSIBILITY BOND", ["50"] = "FINANCING STATEMENT NON STANDARD", ["51"] = "FINANCING STATEMENT STANDARD",
        ["53"] = "FIXTURE FILING NON STANDARD", ["54"] = "FIXTURE FILING STANDARD", ["127"] = "FKA", ["135"] = "FX/F", ["150"] = "GENERAL",
        ["57"] = "LAND CONTRACT", ["58"] = "LEASE", ["59"] = "LIEN", ["149"] = "LIMITED", ["61"] = "MECHANIC'S LIEN", ["62"] = "MERGER",
        ["64"] = "MISCELLANEOUS", ["63"] = "MOBILE HOME LIEN", ["65"] = "MORTGAGE", ["14"] = "MORTGAGE AGREEMENT", ["60"] = "MORTGAGE CORRECTION",
        ["132"] = "NKA", ["66"] = "NOTICE OF COMMENCEMENT", ["67"] = "NOTICE TO COMMENCE SUIT", ["172"] = "OBROW", ["69"] = "OPTION", ["173"] = "OTTO",
        ["71"] = "PARTIAL ASSIGNMENT OF MORTGAGE", ["80"] = "PARTIAL RELEASE FEDERAL LIEN", ["73"] = "PARTIAL RELEASE FINANCING STMT NON STAND",
        ["74"] = "PARTIAL RELEASE FINANCING STMT STANDARD", ["75"] = "PARTIAL RELEASE FIXTURE FILING NON STAND",
        ["76"] = "PARTIAL RELEASE FIXTURE FILING STANDARD", ["79"] = "PARTIAL RELEASE OF MORTGAGE", ["72"] = "PARTNERSHIP",
        ["81"] = "PERSONAL TAX LIEN", ["82"] = "PLANNED UNIT DEVELOPMENT", ["77"] = "PLAT", ["70"] = "POWER OF ATTORNEY", ["133"] = "R/DA",
        ["83"] = "RELEASE AGRICULTURAL LIEN", ["93"] = "RELEASE BY COURT ENTRY", ["102"] = "RELEASE CHILD SUPPORT LIEN",
        ["100"] = "RELEASE MOBILE HOME LIEN", ["91"] = "RELEASE OF ASSIGNMENT OF RENTS", ["92"] = "RELEASE OF BOND LIEN",
        ["84"] = "RELEASE OF EASEMENT", ["94"] = "RELEASE OF EPA LIEN", ["85"] = "RELEASE OF FEDERAL LIEN", ["97"] = "RELEASE OF FEDERAL TAX LIEN",
        ["98"] = "RELEASE OF LAND CONTRACT", ["86"] = "RELEASE OF LEASE", ["87"] = "RELEASE OF LIEN", ["99"] = "RELEASE OF MECHANIC'S LIEN",
        ["88"] = "RELEASE OF MORTGAGE", ["101"] = "RELEASE OF PERSONAL TAX LIEN", ["89"] = "RELEASE OF POWER OF ATTORNEY",
        ["140"] = "RELEASE OF TAX LIEN CERTIFICATE", ["103"] = "RELEASE OF UNEMPLOYMENT LIEN", ["90"] = "RELEASE OF WORKER'S COMPENSATION LIEN",
        ["95"] = "RESOLUTION", ["105"] = "SERVICE DISCHARGE", ["104"] = "SERVICE DISCHARGE COPY", ["106"] = "SHERIFF'S DEED",
        ["111"] = "SPECIAL INSTRUMENT (DEED)", ["112"] = "SPECIAL INSTRUMENT (MORTGAGE)", ["107"] = "SUBORDINATION FINANCING STMT NON STAND",
        ["108"] = "SUBORDINATION FINANCING STMT STAND", ["109"] = "SUBORDINATION FIXTURE FILING NON STAND",
        ["110"] = "SUBORDINATION FIXTURE FILING STANDARD", ["116"] = "SUBORDINATION OF MORTGAGE", ["134"] = "SURV",
        ["139"] = "TAX LIEN CERTIFICATE", ["117"] = "TERMINATION", ["48"] = "UCC FINDING FEE", ["118"] = "UNEMPLOYMENT LIEN",
        ["148"] = "UNKNOWN", ["119"] = "VACATION", ["121"] = "WAIVER", ["174"] = "WELLSTED", ["122"] = "WORKER'S COMPENSATION LIEN",
        ["123"] = "ZONING DOCUMENT", ["125"] = "ZONING DOCUMENTS", ["124"] = "ZONING MAP", ["126"] = "ZONING RESOLUTION"
    };

    /// <summary>By Type form: Sort uses different RISS values than By Name.</summary>
    static readonly Dictionary<string, string> SortOrderToRissByType = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "[FileDateSort] ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "[FileDateSort] DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    /// <summary>By Type form: Document Type*, Include Federal Lien, Index Type*, LastName, FirstName, Side, StartDate, EndDate, Sort.</summary>
    static async Task FillByTypeForm(IPage page, ActorInput input)
    {
        var docTypes = input.DocumentTypes ?? (input.DocumentType != null ? new[] { input.DocumentType } : Array.Empty<string>());
        docTypes = docTypes.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t!.Trim()).ToArray();
        if (docTypes.Length == 0)
            throw new InvalidOperationException("By Type requires at least one document type (e.g. DEED, MORTGAGE).");

        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByType.TryGetValue(sortVal, out var sortOpt) ? sortOpt : "[FileDateSort] ASC";
        var sideVal = !string.IsNullOrWhiteSpace(input.Side) && SideToRissByName.TryGetValue(input.Side.Trim(), out var s) ? s : "%";
        var lastName = input.LastName?.Trim().ToUpperInvariant() ?? "";
        var firstName = input.FirstName?.Trim().ToUpperInvariant() ?? "";
        var lastNameSearch = NormalizeNameSearch(input.LastNameSearch);
        var firstNameSearch = NormalizeNameSearch(input.FirstNameSearch);

        Console.WriteLine($"[FillByType] DocumentTypes=[{string.Join(",", docTypes)}], IncludeFederalLien={input.IncludeFederalLien}, IndexTypes=[{string.Join(",", rissIndex)}], LastName={lastName}, FirstName={firstName}, Side={sideVal}, StartDate={dateVal}, EndDate={endVal}, Sort={rissSort}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [docTypeNames, includeFederal, idxArr, dateVal, endVal, sortVal, sideVal, lastName, firstName, lastNameSearch, firstNameSearch] = args;
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const chosenUpdate2 = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const result = {};

            const docTypeEl = document.querySelector('#document_type, select[name=""selectedtype""]');
            if (docTypeEl && docTypeNames && docTypeNames.length) {
                const targets = docTypeNames.map(n => (n || '').trim()).filter(Boolean);
                const targetsUpper = targets.map(t => t.toUpperCase());
                Array.from(docTypeEl.options).forEach(o => {
                    const val = (o.value || '').trim();
                    const txt = (o.textContent || '').trim();
                    const txtUpper = txt.toUpperCase();
                    o.selected = targets.some((t, i) => {
                        const tu = (t || '').toUpperCase();
                        return val === t || txt === t || txtUpper === tu || (tu.length >= 3 && txtUpper.includes(tu));
                    });
                });
                fireChange(docTypeEl);
                if (typeof jQuery !== 'undefined' && jQuery(docTypeEl).data('chosen')) jQuery(docTypeEl).trigger('chosen:updated');
                result.docTypeSelected = Array.from(docTypeEl.options).filter(o => o.selected).map(o => o.value);
            }

            const fedEl = document.querySelector('#federallien, input[name=""federallien""]');
            if (fedEl) { fedEl.checked = !!includeFederal; fireChange(fedEl); result.includeFederal = fedEl.checked; }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate2(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const lastEl = document.querySelector('#LastName, input[name=""LastName""]');
            const firstEl = document.querySelector('#FirstName, input[name=""FirstName""]');
            const lastSearchEl = document.querySelector('#LastName_Search, select[name=""LastName_Search""]');
            const firstSearchEl = document.querySelector('#FirstName_Search, select[name=""FirstName_Search""]');
            const startEl = document.querySelector('#StartDate, input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDate, input[name=""EndDate""]');
            const sideEl = document.querySelector('#Side, select[name=""Side""]');
            const sortEl = document.querySelector('#Sort, select[name=""sort""]');
            if (lastEl) { lastEl.value = lastName || ''; fireChange(lastEl); result.lastName = lastEl.value; }
            if (firstEl) { firstEl.value = firstName || ''; fireChange(firstEl); }
            if (lastSearchEl && lastNameSearch) { lastSearchEl.value = lastNameSearch; fireChange(lastSearchEl); chosenUpdate(lastSearchEl); }
            if (firstSearchEl && firstNameSearch) { firstSearchEl.value = firstNameSearch; fireChange(firstSearchEl); chosenUpdate(firstSearchEl); }
            if (startEl) { startEl.value = dateVal || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = endVal || ''; fireChange(endEl); result.endValue = endEl.value; }
            if (sideEl && sideVal) { sideEl.value = sideVal; fireChange(sideEl); chosenUpdate(sideEl); result.sideValue = sideEl.value; }
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }
            return result;
        }", new object[] { docTypes, input.IncludeFederalLien, rissIndex, dateVal, endVal, rissSort, sideVal, lastName, firstName, lastNameSearch, firstNameSearch });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByType] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByType] Could not log result: {ex.Message}"); }
    }

    /// <summary>By Municipality: Properties*, IndexType*, Lot, Sort, More Filters: LastName, FirstName, LastName_Search, FirstName_Search, StartDate, EndDate.</summary>
    static async Task FillByMunicipalityForm(IPage page, ActorInput input)
    {
        var propertiesVal = (input.Properties ?? input.Municipality)?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(propertiesVal))
            throw new InvalidOperationException("Properties is required for 'By Municipality' search (e.g. DAYTON, BACHMAN).");

        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var lot = input.Lot?.Trim() ?? "";
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByName.TryGetValue(sortVal, out var sortOpt) ? sortOpt : "CAST([FileDate] as DATE) ASC";
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var lastName = input.LastName?.Trim().ToUpperInvariant() ?? "";
        var firstName = input.FirstName?.Trim().ToUpperInvariant() ?? "";
        var lastNameSearch = NormalizeNameSearch(input.LastNameSearch);
        var firstNameSearch = NormalizeNameSearch(input.FirstNameSearch);

        Console.WriteLine($"[FillByMunicipality] Properties={propertiesVal}, Lot={lot}, IndexTypes=[{string.Join(",", rissIndex)}], Sort={rissSort}, LastName={lastName}, FirstName={firstName}");

        var moreFiltersBtn = page.Locator("button.togglebutton[data-id='addition'], button:has-text('More Filters')").First;
        if (await moreFiltersBtn.CountAsync() > 0)
        {
            try { await moreFiltersBtn.ClickAsync(); await page.WaitForTimeoutAsync(300); } catch { }
        }

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [propertiesVal, lot, idxArr, sortVal, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch] = args;
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const result = {};

            const propEl = document.querySelector('select[name=""properties""], #properties');
            if (propEl && propertiesVal) {
                const mun = (propertiesVal || '').trim().toUpperCase();
                const opt = Array.from(propEl.options).find(o => (o.value || '').toUpperCase() === mun || (o.textContent || '').trim().toUpperCase() === mun);
                if (opt) { propEl.value = opt.value; fireChange(propEl); chosenUpdate(propEl); result.properties = propEl.value; }
            }

            const lotEl = document.querySelector('#lot, input[name=""lot""]');
            if (lotEl) { lotEl.value = lot || ''; fireChange(lotEl); result.lot = lotEl.value; }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }

            const lastEl = document.querySelector('#LastName, input[name=""LastName""]');
            const firstEl = document.querySelector('#FirstName, input[name=""FirstName""]');
            const lastSearchEl = document.querySelector('#LastName_Search, select[name=""LastName_Search""]');
            const firstSearchEl = document.querySelector('#FirstName_Search, select[name=""FirstName_Search""]');
            const startEl = document.querySelector('#StartDate, input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDate, input[name=""EndDate""]');
            if (lastEl) { lastEl.value = lastName || ''; fireChange(lastEl); result.lastName = lastEl.value; }
            if (firstEl) { firstEl.value = firstName || ''; fireChange(firstEl); }
            if (lastSearchEl && lastNameSearch) { lastSearchEl.value = lastNameSearch; fireChange(lastSearchEl); }
            if (firstSearchEl && firstNameSearch) { firstSearchEl.value = firstNameSearch; fireChange(firstSearchEl); }
            if (startEl) { startEl.value = dateVal || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = endVal || ''; fireChange(endEl); result.endValue = endEl.value; }
            return result;
        }", new object[] { propertiesVal, lot, rissIndex, rissSort, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByMunicipality] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByMunicipality] Could not log result: {ex.Message}"); }
    }

    /// <summary>By Subdivision: Properties* (multi-select in modal #getProperties), IndexType*, Lot, Sort, More Filters.</summary>
    static async Task FillBySubdivisionForm(IPage page, ActorInput input)
    {
        var subs = input.Subdivisions ?? (input.Subdivision != null ? new[] { input.Subdivision } : Array.Empty<string>());
        subs = subs.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t!.Trim()).ToArray();
        if (subs.Length == 0)
            throw new InvalidOperationException("By Subdivision requires at least one property (e.g. subdivisions: [\"ABBEY (LOTS 9 - 32)\"]).");

        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var lot = input.Lot?.Trim() ?? "";
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByName.TryGetValue(sortVal, out var sortOpt) ? sortOpt : "CAST([FileDate] as DATE) ASC";
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var lastName = input.LastName?.Trim().ToUpperInvariant() ?? "";
        var firstName = input.FirstName?.Trim().ToUpperInvariant() ?? "";
        var lastNameSearch = NormalizeNameSearch(input.LastNameSearch);
        var firstNameSearch = NormalizeNameSearch(input.FirstNameSearch);

        Console.WriteLine($"[FillBySubdivision] Subdivisions=[{string.Join(",", subs)}], Lot={lot}, IndexTypes=[{string.Join(",", rissIndex)}], Sort={rissSort}");

        var moreFiltersBtn = page.Locator("button.togglebutton[data-id='addition'], button:has-text('More Filters')").First;
        if (await moreFiltersBtn.CountAsync() > 0)
        {
            try { await moreFiltersBtn.ClickAsync(); await page.WaitForTimeoutAsync(300); } catch { }
        }

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [subNames, lot, idxArr, sortVal, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch] = args;
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const result = {};

            const propEl = document.querySelector('#properties, select[name=""properties""]');
            if (propEl && subNames && subNames.length) {
                const targets = subNames.map(s => (s || '').trim().toUpperCase());
                const targetsUpper = targets;
                Array.from(propEl.options).forEach(o => {
                    const val = (o.value || '').trim();
                    const txt = (o.textContent || '').trim();
                    const txtUpper = txt.toUpperCase();
                    const match = targets.some((t, i) => {
                        const tu = (t || '').toUpperCase();
                        return val === t || txt === t || txtUpper === tu || (tu.length >= 3 && txtUpper.includes(tu));
                    });
                    o.selected = match;
                });
                fireChange(propEl);
                result.propertiesSelected = Array.from(propEl.options).filter(o => o.selected).map(o => o.value);
            }

            const lotEl = document.querySelector('#lot, input[name=""lot""]');
            if (lotEl) { lotEl.value = lot || ''; fireChange(lotEl); result.lot = lotEl.value; }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }

            const lastEl = document.querySelector('#LastName, input[name=""LastName""]');
            const firstEl = document.querySelector('#FirstName, input[name=""FirstName""]');
            const lastSearchEl = document.querySelector('#LastName_Search, select[name=""LastName_Search""]');
            const firstSearchEl = document.querySelector('#FirstName_Search, select[name=""FirstName_Search""]');
            const startEl = document.querySelector('#StartDate, input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDate, input[name=""EndDate""]');
            if (lastEl) { lastEl.value = lastName || ''; fireChange(lastEl); result.lastName = lastEl.value; }
            if (firstEl) { firstEl.value = firstName || ''; fireChange(firstEl); }
            if (lastSearchEl && lastNameSearch) { lastSearchEl.value = lastNameSearch; fireChange(lastSearchEl); }
            if (firstSearchEl && firstNameSearch) { firstSearchEl.value = firstNameSearch; fireChange(firstSearchEl); }
            if (startEl) { startEl.value = dateVal || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = endVal || ''; fireChange(endEl); result.endValue = endEl.value; }
            return result;
        }", new object[] { subs, lot, rissIndex, rissSort, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillBySubdivision] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillBySubdivision] Could not log result: {ex.Message}"); }
    }

    /// <summary>By STR: Properties*, IndexType*, Section, Township, Range, Sort, More Filters.</summary>
    static async Task FillBySTRForm(IPage page, ActorInput input)
    {
        var propertiesVal = (input.Properties ?? input.Municipality ?? "ALL")?.Trim() ?? "ALL";
        if (string.IsNullOrWhiteSpace(propertiesVal)) propertiesVal = "ALL";

        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var sectionVal = input.Section?.Trim() ?? "";
        var townshipVal = input.Township?.Trim() ?? "";
        var rangeVal = input.Range?.Trim() ?? "";
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByName.TryGetValue(sortVal, out var sortOpt) ? sortOpt : "CAST([FileDate] as DATE) ASC";
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var lastName = input.LastName?.Trim().ToUpperInvariant() ?? "";
        var firstName = input.FirstName?.Trim().ToUpperInvariant() ?? "";
        var lastNameSearch = NormalizeNameSearch(input.LastNameSearch);
        var firstNameSearch = NormalizeNameSearch(input.FirstNameSearch);

        Console.WriteLine($"[FillBySTR] Properties={propertiesVal}, Section={sectionVal}, Township={townshipVal}, Range={rangeVal}, IndexTypes=[{string.Join(",", rissIndex)}], Sort={rissSort}");

        var moreFiltersBtn = page.Locator("button.togglebutton[data-id='addition'], button:has-text('More Filters')").First;
        if (await moreFiltersBtn.CountAsync() > 0)
        {
            try { await moreFiltersBtn.ClickAsync(); await page.WaitForTimeoutAsync(300); } catch { }
        }

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [propertiesVal, sectionVal, townshipVal, rangeVal, idxArr, sortVal, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch] = args;
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };
            const matchOpt = (opts, userVal) => {
                if (!userVal) return null;
                const u = (userVal || '').trim();
                const opt = Array.from(opts).find(o => (o.value || '').trim() === u || (o.value || '').trim().toUpperCase() === u.toUpperCase());
                return opt ? opt.value : null;
            };
            const result = {};

            const propEl = document.querySelector('select[name=""properties""], #properties');
            if (propEl && propertiesVal) {
                const p = (propertiesVal || '').toUpperCase().trim();
                const opt = Array.from(propEl.options).find(o => (o.value || '').trim().toUpperCase() === p || (o.textContent || '').trim().toUpperCase() === p);
                if (opt) { propEl.value = opt.value; fireChange(propEl); chosenUpdate(propEl); result.properties = propEl.value; }
            }

            const sectEl = document.querySelector('#section, select[name=""section""]');
            if (sectEl && sectionVal) {
                const v = matchOpt(sectEl.options, sectionVal);
                if (v) { sectEl.value = v; fireChange(sectEl); chosenUpdate(sectEl); result.section = sectEl.value; }
            }

            const townEl = document.querySelector('#township, select[name=""township""]');
            if (townEl && townshipVal) {
                const v = matchOpt(townEl.options, townshipVal);
                if (v) { townEl.value = v; fireChange(townEl); chosenUpdate(townEl); result.township = townEl.value; }
            }

            const rangeEl = document.querySelector('#range, select[name=""range""]');
            if (rangeEl && rangeVal) {
                const v = matchOpt(rangeEl.options, rangeVal);
                if (v) { rangeEl.value = v; fireChange(rangeEl); chosenUpdate(rangeEl); result.range = rangeEl.value; }
            }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }

            const lastEl = document.querySelector('#LastName, input[name=""LastName""]');
            const firstEl = document.querySelector('#FirstName, input[name=""FirstName""]');
            const lastSearchEl = document.querySelector('#LastName_Search, select[name=""LastName_Search""]');
            const firstSearchEl = document.querySelector('#FirstName_Search, select[name=""FirstName_Search""]');
            const startEl = document.querySelector('#StartDate, input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDate, input[name=""EndDate""]');
            if (lastEl) { lastEl.value = lastName || ''; fireChange(lastEl); result.lastName = lastEl.value; }
            if (firstEl) { firstEl.value = firstName || ''; fireChange(firstEl); }
            if (lastSearchEl && lastNameSearch) { lastSearchEl.value = lastNameSearch; fireChange(lastSearchEl); }
            if (firstSearchEl && firstNameSearch) { firstSearchEl.value = firstNameSearch; fireChange(firstSearchEl); }
            if (startEl) { startEl.value = dateVal || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = endVal || ''; fireChange(endEl); result.endValue = endEl.value; }
            return result;
        }", new object[] { propertiesVal, sectionVal, townshipVal, rangeVal, rissIndex, rissSort, dateVal, endVal, lastName, firstName, lastNameSearch, firstNameSearch });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillBySTR] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillBySTR] Could not log result: {ex.Message}"); }
    }

    /// <summary>By Instrument: InstrumentYear*, InstrumentNumber*, IndexType* (multi), Sort, Captcha.</summary>
    static async Task FillByInstrumentForm(IPage page, ActorInput input)
    {
        var yearVal = (input.InstrumentYear ?? "").Trim();
        var numVal = (input.InstrumentNumber ?? "").Trim();
        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRiss.TryGetValue(sortVal, out var s) ? s : "[FileDateSort] ASC";

        Console.WriteLine($"[FillByInstrument] Year={yearVal}, Number={numVal}, IndexTypes=[{string.Join(",", rissIndex)}], Sort={rissSort}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [yearVal, numVal, idxArr, sortVal] = args;
            const result = {};
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };

            const yearEl = document.querySelector('#InstrumentYear, input[name=""InstrumentYear""]');
            if (yearEl && yearVal) { yearEl.value = yearVal; fireChange(yearEl); result.year = yearEl.value; }

            const numEl = document.querySelector('#InstrumentNumber, input[name=""InstrumentNumber""]');
            if (numEl && numVal) { numEl.value = numVal; fireChange(numEl); result.number = numEl.value; }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }

            return result;
        }", new object[] { yearVal, numVal, rissIndex, rissSort });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByInstrument] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByInstrument] Could not log result: {ex.Message}"); }
    }

    static readonly Dictionary<string, string> IndexTypeToRissBookPage = new(StringComparer.OrdinalIgnoreCase)
    {
        ["All"] = "%", ["%"] = "%",
        ["DEED"] = "DBK", ["DBK"] = "DBK", ["Deeds"] = "DBK",
        ["MTG"] = "MBK", ["MBK"] = "MBK", ["Mortgages"] = "MBK", ["Mortgage"] = "MBK",
        ["PLT"] = "PBK", ["PBK"] = "PBK"
    };

    /// <summary>By Book/Page: Book*, Page*, Thru (optional), IndexType (All/DEED/MTG/PLT).</summary>
    static async Task FillByBookPageForm(IPage page, ActorInput input)
    {
        var bookVal = (input.Book ?? "").Trim();
        var pageVal = (input.Page ?? "").Trim();
        var thruVal = (input.PageThru ?? "").Trim();
        var idxVal = (input.BookPageIndexType ?? "All").Trim();
        var rissIdx = !string.IsNullOrEmpty(idxVal) && IndexTypeToRissBookPage.TryGetValue(idxVal, out var v) ? v : "%";

        Console.WriteLine($"[FillByBookPage] Book={bookVal}, Page={pageVal}, Thru={thruVal}, IndexType={rissIdx}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [bookVal, pageVal, thruVal, idxVal] = args;
            const result = {};
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };

            const bookEl = document.querySelector('#Book, input[name=""Book""]');
            if (bookEl && bookVal) { bookEl.value = bookVal; fireChange(bookEl); result.book = bookEl.value; }

            const pageEl = document.querySelector('#Page, input[name=""Page""]');
            if (pageEl && pageVal) { pageEl.value = pageVal; fireChange(pageEl); result.page = pageEl.value; }

            const thruEl = document.querySelector('#Thru, input[name=""Thru""]');
            if (thruEl && thruVal) { thruEl.value = thruVal; fireChange(thruEl); result.thru = thruEl.value; }

            const idxEl = document.querySelector('#IndexTypeBP, select[name=""IndexType""]');
            if (idxEl && idxVal) {
                const opt = Array.from(idxEl.options).find(o => (o.value || '') === idxVal);
                if (opt) { idxEl.value = idxVal; fireChange(idxEl); chosenUpdate(idxEl); result.indexType = idxEl.value; }
            }
            return result;
        }", new object[] { bookVal, pageVal, thruVal, rissIdx });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByBookPage] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByBookPage] Could not log result: {ex.Message}"); }
    }

    /// <summary>By Fiche: Fiche* (6-10 alphanumeric), IndexType* (multi), Sort (FILE DATE ASC/DESC).</summary>
    static async Task FillByFicheForm(IPage page, ActorInput input)
    {
        var ficheVal = (input.Fiche ?? "").Trim();
        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        indexTypes = indexTypes.Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissIndex.Length == 0) rissIndex = ["DEE", "MTG"];

        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRissByName.TryGetValue(sortVal, out var s) ? s : "CAST([FileDate] as DATE) DESC";

        Console.WriteLine($"[FillByFiche] Fiche={ficheVal}, IndexTypes=[{string.Join(",", rissIndex)}], Sort={rissSort}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [ficheVal, idxArr, sortVal] = args;
            const result = {};
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };

            const ficheEl = document.querySelector('#Fiche, input[name=""Fiche""]');
            if (ficheEl && ficheVal) { ficheEl.value = ficheVal; fireChange(ficheEl); result.fiche = ficheEl.value; }

            const idxEl = document.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxArr && idxArr.length) {
                Array.from(idxEl.options).forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl); chosenUpdate(idxEl);
                result.idxSelected = Array.from(idxEl.options).filter(o => o.selected).map(o => o.value);
            }

            const sortEl = document.querySelector('#Sort, select[name=""Sort""]');
            if (sortEl && sortVal) { sortEl.value = sortVal; fireChange(sortEl); chosenUpdate(sortEl); result.sortValue = sortEl.value; }

            return result;
        }", new object[] { ficheVal, rissIndex, rissSort });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByFiche] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByFiche] Could not log result: {ex.Message}"); }
    }

    static readonly Dictionary<string, string> IndexTypeToRissPre1980 = new(StringComparer.OrdinalIgnoreCase)
    {
        ["All"] = "%", ["%"] = "%",
        ["PDE"] = "PDE", ["DEED"] = "PDE", ["Deeds"] = "PDE",
        ["PMT"] = "PMT", ["MTG"] = "PMT", ["Mortgages"] = "PMT", ["Mortgage"] = "PMT"
    };

    /// <summary>By Pre-1980: PreYear (1971-1979), PreNumber* (required), IndexType (All/PDE/PMT).</summary>
    static async Task FillByPre1980Form(IPage page, ActorInput input)
    {
        var yearVal = (input.Pre1980Year ?? "1971").Trim();
        if (string.IsNullOrEmpty(yearVal) || !int.TryParse(yearVal, out var y) || y < 1971 || y > 1979)
            yearVal = "1971";
        var numVal = (input.Pre1980Number ?? "").Trim();
        var idxVal = (input.Pre1980IndexType ?? "All").Trim();
        var rissIdx = !string.IsNullOrEmpty(idxVal) && IndexTypeToRissPre1980.TryGetValue(idxVal, out var v) ? v : "%";

        Console.WriteLine($"[FillByPre1980] PreYear={yearVal}, PreNumber={numVal}, IndexType={rissIdx}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [yearVal, numVal, idxVal] = args;
            const form = document.querySelector('#searchFormPre, form[action*=""pre1980_search""]');
            if (!form) return { error: 'Form not found' };
            const result = {};
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            const chosenUpdate = (el) => { if (el && typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated'); };

            const yearEl = form.querySelector('#PreYear, select[name=""PreYear""]');
            if (yearEl && yearVal) {
                const opt = Array.from(yearEl.options).find(o => (o.value || '') === yearVal);
                if (opt) { yearEl.value = yearVal; fireChange(yearEl); chosenUpdate(yearEl); result.preYear = yearEl.value; }
            }

            const numEl = form.querySelector('#PreNumber, input[name=""PreNumber""]');
            if (numEl && numVal) { numEl.value = numVal; fireChange(numEl); result.preNumber = numEl.value; }

            const idxEl = form.querySelector('#IndexType, select[name=""IndexType""]');
            if (idxEl && idxVal) {
                const opt = Array.from(idxEl.options).find(o => (o.value || '') === idxVal);
                if (opt) { idxEl.value = idxVal; fireChange(idxEl); chosenUpdate(idxEl); result.indexType = idxEl.value; }
            }
            return result;
        }", new object[] { yearVal, numVal, rissIdx });

        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByPre1980] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByPre1980] Could not log result: {ex.Message}"); }
    }

    static readonly Dictionary<string, string> IndexTypeToRiss = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Deeds"] = "DEE", ["Mortgages"] = "MTG", ["Partnerships"] = "PTR",
        ["UCC"] = "FIN", ["Service Discharge"] = "DIS", ["Veteran Graves"] = "VET"
    };

    static readonly Dictionary<string, string> SortOrderToRiss = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "[FileDateSort] ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "[FileDateSort] DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    static async Task FillByDateFormBatchAsync(IPage page, ActorInput input, bool fillCaptcha)
    {
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissIndex = indexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t) && IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        var sortVal = input.SortOrder?.Trim();
        var rissSort = !string.IsNullOrEmpty(sortVal) && SortOrderToRiss.TryGetValue(sortVal, out var s) ? s : "[InstrumentNum] ASC";

        Console.WriteLine($"[FillByDate] Filling form: StartDate={dateVal}, EndDate={endVal}, IndexTypes(RISS)=[{string.Join(",", rissIndex)}], Sort={rissSort}");

        var fillResult = await page.EvaluateAsync<object>(@"args => {
            const [d, end, idxArr, sortVal] = args;
            const startEl = document.querySelector('#StartDateD') || document.querySelector('input[name=""StartDate""]');
            const endEl = document.querySelector('#EndDateD') || document.querySelector('input[name=""EndDate""]');
            const idxEl = document.querySelector('#IndexType') || document.querySelector('select[name=""IndexType""]');
            const sortEl = document.querySelector('#Sort') || document.querySelector('select[name=""Sort""]');
            const result = { startFound: !!startEl, endFound: !!endEl, idxFound: !!idxEl, sortFound: !!sortEl };
            const fireChange = (el) => { if (el) { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); } };
            if (startEl) { startEl.value = d || ''; fireChange(startEl); result.startValue = startEl.value; }
            if (endEl) { endEl.value = end || ''; fireChange(endEl); result.endValue = endEl.value; }
            if (idxEl && idxArr && idxArr.length) {
                const opts = Array.from(idxEl.options);
                opts.forEach(o => { o.selected = idxArr.includes(o.value); });
                fireChange(idxEl);
                if (typeof jQuery !== 'undefined' && jQuery(idxEl).data('chosen')) jQuery(idxEl).trigger('chosen:updated');
                result.idxSelected = opts.filter(o => o.selected).map(o => o.value);
            }
            if (sortEl && sortVal) {
                sortEl.value = sortVal;
                sortEl.dispatchEvent(new Event('change', { bubbles: true }));
                if (typeof jQuery !== 'undefined' && jQuery(sortEl).data('chosen')) jQuery(sortEl).trigger('chosen:updated');
                result.sortValue = sortEl.value;
            }
            return result;
        }", new object[] { dateVal, endVal, rissIndex, rissSort });
        try
        {
            var r = fillResult != null ? JsonSerializer.Serialize(fillResult) : "null";
            Console.WriteLine($"[FillByDate] Form fill result: {r}");
        }
        catch (Exception ex) { Console.WriteLine($"[FillByDate] Could not log result: {ex.Message}"); }

        if (fillCaptcha && !string.IsNullOrWhiteSpace(input.CaptchaCode))
        {
            var cap = input.CaptchaCode.Trim().ToUpperInvariant();
            await page.Locator("#captchaid, input[name='captcha']").First.FillAsync(cap);
        }
    }

    static async Task FillCommonFields(IPage page, ActorInput input)
    {
        var dateVal = input.StartDate ?? DateTime.Now.ToString("yyyy-MM-dd");
        var endVal = input.EndDate ?? dateVal;
        await page.Locator("#StartDateD, input[name='StartDate']").First.FillAsync(dateVal);
        await page.Locator("#EndDateD, input[name='EndDate']").First.FillAsync(endVal);

        var indexTypes = input.IndexTypes ?? ["Deeds", "Mortgages"];
        if (indexTypes.Length == 0) indexTypes = ["Deeds", "Mortgages"];
        var rissValues = indexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t) && IndexTypeToRiss.TryGetValue(t.Trim(), out _))
            .Select(t => IndexTypeToRiss[t.Trim()])
            .Distinct()
            .ToArray();
        if (rissValues.Length > 0)
        {
            var sel = page.Locator("select#IndexType, select[name='IndexType']").First;
            if (await sel.CountAsync() > 0)
            {
                try
                {
                    await sel.SelectOptionAsync(rissValues, new LocatorSelectOptionOptions { Force = true });
                }
                catch
                {
                    await sel.EvaluateAsync(@"([el, vals]) => {
                        Array.from(el.options).forEach(o => { o.selected = vals.includes(o.value); });
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        if (typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated');
                    }", rissValues);
                }
            }
        }

        var sortVal = input.SortOrder?.Trim();
        if (!string.IsNullOrEmpty(sortVal) && SortOrderToRiss.TryGetValue(sortVal, out var rissSort))
            await TrySelectChosenAsync(page, "Sort", rissSort);
    }

    static async Task TrySelectChosenAsync(IPage page, string labelOrId, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var sel = page.GetByLabel(labelOrId).Or(page.Locator("select#Sort, select[name='Sort'], select.chosen-select")).First;
        if (await sel.CountAsync() == 0) return;
        try
        {
            await sel.First.SelectOptionAsync(new[] { value }, new LocatorSelectOptionOptions { Force = true });
        }
        catch
        {
            try
            {
                await sel.First.EvaluateAsync(@"el => {
                    el.value = arguments[1];
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof jQuery !== 'undefined' && jQuery(el).data('chosen')) jQuery(el).trigger('chosen:updated');
                }", value);
            }
            catch { }
        }
    }

    static async Task TryFillAsync(IPage page, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var el = page.GetByLabel(label).Or(page.Locator($"input[id*='{label.Replace(" ", "")}'], input[name*='{label.Replace(" ", "")}']").First);
        if (await el.CountAsync() > 0)
            await el.First.FillAsync(value);
    }

    static async Task TrySelectAsync(IPage page, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var sel = page.GetByLabel(label).Or(page.Locator("select[id*='Side'], select[name*='Side']").First);
        if (await sel.CountAsync() > 0)
            await sel.First.SelectOptionAsync(new[] { value });
    }

    static async Task TryFillOrSelectAsync(IPage page, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var lbl = page.GetByLabel(label);
        if (await lbl.CountAsync() > 0)
        {
            var tag = await lbl.First.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
            if (tag == "select")
                await lbl.First.SelectOptionAsync(new[] { value });
            else
                await lbl.First.FillAsync(value);
        }
    }

    static async Task<string> GetDetailValue(IPage page, string label)
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

    static async Task<string> GetDetailValueList(IPage page, string label)
    {
        try
        {
            var group = page.Locator("div.input-group").Filter(new() { Has = page.Locator($"span.informationTitle:has-text('{label}')") }).First;
            if (await group.CountAsync() == 0) return "";
            var dataEl = group.Locator(".informationData");
            if (await dataEl.CountAsync() == 0) return "";
            var parts = new List<string>();
            var childTags = dataEl.Locator("p, div");
            var n = await childTags.CountAsync();
            if (n > 0)
            {
                for (var i = 0; i < n; i++)
                {
                    var t = (await childTags.Nth(i).InnerTextAsync()).Trim();
                    if (!string.IsNullOrEmpty(t)) parts.Add(t);
                }
            }
            else
            {
                var t = (await dataEl.InnerTextAsync()).Trim();
                if (!string.IsNullOrEmpty(t)) parts.Add(t);
            }
            return string.Join(";", parts.Where(s => !string.IsNullOrWhiteSpace(s)));
        }
        catch { }
        return "";
    }

    static async Task<string> GetRemarksList(IPage page)
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

    static async Task<(IBrowserContext ctx, IPage page, int rowCount, int fileNumberColIndex, int fileDateColIndex, int imageColIndex, string tableSelector, bool isByName)> CreateContextAndLoadResultsAsync(
        IBrowser browser, ActorInput input, bool needsCaptcha)
    {
        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            ViewportSize = new ViewportSize { Width = 1024, Height = 768 },
            IgnoreHTTPSErrors = true,
            Permissions = new[] { "clipboard-read", "clipboard-write" }
        });
        context.SetDefaultTimeout(60000);
        context.SetDefaultNavigationTimeout(90000);

        void AttachDialogHandler(IPage p) => p.Dialog += async (_, d) => { try { await d.AcceptAsync(); } catch { } };
        context.Page += (_, np) => AttachDialogHandler(np);

        var page = await context.NewPageAsync();
        AttachDialogHandler(page);
        page.SetDefaultTimeout(60000);

        await page.GotoAsync("https://riss.mcrecorder.org", new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 120000 });
        var disclaimerCheckbox = page.GetByLabel("I AGREE TO THE LEGAL DISCLAIMER").Or(page.Locator("input[type='checkbox']").First);
        await disclaimerCheckbox.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        await disclaimerCheckbox.CheckAsync();
        await page.WaitForTimeoutAsync(1000);

        var guestButton = page.GetByRole(AriaRole.Button, new() { Name = "Proceed to RISS as GUEST", Exact = true })
            .Or(page.Locator("button:has-text('Proceed to RISS as GUEST')"))
            .Or(page.Locator("input[value='Proceed to RISS as GUEST']"));
        await guestButton.First.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        await guestButton.First.ClickAsync();
        await page.WaitForTimeoutAsync(2000);

        var searchTab = page.Locator("text=Search Public Records").First;
        if (await searchTab.CountAsync() > 0)
        {
            await searchTab.ClickAsync();
            await page.WaitForTimeoutAsync(3000); // Allow sidebar tabs to fully render
        }

        await ClickSidebarTabAsync(page, input.SearchMode);
        await page.WaitForTimeoutAsync(1500);

        // Ensure correct form is visible
        if (input.SearchMode.Trim().Equals("ByDate", StringComparison.OrdinalIgnoreCase))
        {
            var byDateField = page.Locator("#StartDateD, #StartDate, input[name='StartDate'], input[type='date']").First;
            await byDateField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByName", StringComparison.OrdinalIgnoreCase))
        {
            var byNameField = page.Locator("#LastName, input[name='LastName']").First;
            await byNameField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByType", StringComparison.OrdinalIgnoreCase))
        {
            var byTypeField = page.Locator("#document_type, select[name='selectedtype']").First;
            await byTypeField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByMunicipality", StringComparison.OrdinalIgnoreCase))
        {
            var byMunField = page.Locator("select[name='properties'], #lot").First;
            await byMunField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("BySubdivision", StringComparison.OrdinalIgnoreCase))
        {
            var bySubField = page.Locator("#properties, #IndexType, #lot").First;
            await bySubField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("BySTR", StringComparison.OrdinalIgnoreCase))
        {
            var byStrField = page.Locator("select[name='properties'], #section, #IndexType").First;
            await byStrField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByInstrument", StringComparison.OrdinalIgnoreCase))
        {
            var byInstField = page.Locator("#InstrumentYear, #InstrumentNumber, input[name='InstrumentYear']").First;
            await byInstField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByBookPage", StringComparison.OrdinalIgnoreCase))
        {
            var byBookPageField = page.Locator("#Book, #Page, input[name='Book']").First;
            await byBookPageField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByFiche", StringComparison.OrdinalIgnoreCase))
        {
            var byFicheField = page.Locator("#Fiche, input[name='Fiche']").First;
            await byFicheField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }
        else if (input.SearchMode.Trim().Equals("ByPre1980", StringComparison.OrdinalIgnoreCase))
        {
            var byPreField = page.Locator("#PreYear, #PreNumber, input[name='PreNumber']").First;
            await byPreField.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Attached, Timeout = 15000 });
        }

        if (needsCaptcha)
        {
            var captchaBase64 = await ExtractCaptchaBase64Async(page);
            var apiKey = input.TwoCaptchaApiKey?.Trim() ?? Environment.GetEnvironmentVariable("TWO_CAPTCHA_API_KEY")?.Trim();
            if (!string.IsNullOrEmpty(apiKey) && !string.IsNullOrEmpty(captchaBase64))
            {
                var solved = await TwoCaptchaService.SolveImageCaptchaAsync(captchaBase64, apiKey);
                input.CaptchaCode = solved ?? "";
                if (string.IsNullOrEmpty(solved))
                    Console.WriteLine("[2Captcha] Solving failed. Retrying may help.");
            }
            if (string.IsNullOrEmpty(input.CaptchaCode))
                throw new InvalidOperationException("Captcha required but could not be solved. Ensure twoCaptchaApiKey is valid and has credits.");
        }

        switch (input.SearchMode.Trim())
        {
            case "ByDate": await FillByDateFormBatchAsync(page, input, needsCaptcha); break;
            case "ByName": await FillByNameFormBatchAsync(page, input); break;
            case "ByType": await FillByTypeForm(page, input); break;
            case "ByMunicipality": await FillByMunicipalityForm(page, input); break;
            case "BySubdivision": await FillBySubdivisionForm(page, input); break;
            case "BySTR": await FillBySTRForm(page, input); break;
            case "ByInstrument":
                await FillByInstrumentForm(page, input);
                if (needsCaptcha && !string.IsNullOrWhiteSpace(input.CaptchaCode)) await FillCaptchaAsync(page, input.CaptchaCode);
                break;
            case "ByBookPage": await FillByBookPageForm(page, input); break;
            case "ByFiche": await FillByFicheForm(page, input); break;
            case "ByPre1980": await FillByPre1980Form(page, input); break;
            default: throw new InvalidOperationException($"Unknown search mode: {input.SearchMode}");
        }

        var isByName = input.SearchMode.Trim().Equals("ByName", StringComparison.OrdinalIgnoreCase);
        var isByType = input.SearchMode.Trim().Equals("ByType", StringComparison.OrdinalIgnoreCase);
        var isByMunicipality = input.SearchMode.Trim().Equals("ByMunicipality", StringComparison.OrdinalIgnoreCase);
        var isBySubdivision = input.SearchMode.Trim().Equals("BySubdivision", StringComparison.OrdinalIgnoreCase);
        var isBySTR = input.SearchMode.Trim().Equals("BySTR", StringComparison.OrdinalIgnoreCase);
        var isByInstrument = input.SearchMode.Trim().Equals("ByInstrument", StringComparison.OrdinalIgnoreCase);
        var isByBookPage = input.SearchMode.Trim().Equals("ByBookPage", StringComparison.OrdinalIgnoreCase);
        var isByFiche = input.SearchMode.Trim().Equals("ByFiche", StringComparison.OrdinalIgnoreCase);
        var isByPre1980 = input.SearchMode.Trim().Equals("ByPre1980", StringComparison.OrdinalIgnoreCase);
        var isFormSubmitMode = isByName || isByType || isByMunicipality || isBySubdivision || isBySTR || isByInstrument || isByBookPage || isByFiche || isByPre1980;
        Console.WriteLine($"[Search] Submitting {input.SearchMode} form...");
        if (isByName)
        {
            await page.EvaluateAsync(@"() => {
                const form = document.querySelector('#myForm1') || document.querySelector('form:has(#LastName)') || document.querySelector('form');
                if (form) { form.action = '/name_search/all_matches.cfm'; form.submit(); }
            }");
        }
        else if (isByType)
        {
            await page.EvaluateAsync(@"() => {
                const form = document.querySelector('#myForm1') || document.querySelector('form:has(#document_type)') || document.querySelector('form');
                if (form) { form.action = '/type_search/all_matches.cfm'; form.submit(); }
            }");
        }
        else if (isByMunicipality || isBySubdivision)
        {
            var searchBtn = page.Locator("#searchP").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            else
                await page.EvaluateAsync(@"() => { const f = document.querySelector('#myForm'); if (f) f.submit(); }");
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        else if (isBySTR || isByFiche)
        {
            var searchBtn = page.Locator("#search").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            else
            {
                var formSelector = isByFiche ? "form[action*='fiche_search']" : "form[action*='str_search']";
                await page.EvaluateAsync($@"() => {{ const f = document.querySelector('{formSelector}') || document.querySelector('#myForm'); if (f) f.submit(); }}");
            }
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        else if (isByInstrument)
        {
            var searchBtn = page.Locator("#searchI").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            else
                await page.EvaluateAsync(@"() => { const f = document.querySelector('form[name=""searchForm3""]'); if (f) f.submit(); }");
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        else if (isByBookPage)
        {
            var searchBtn = page.Locator("#searchBP").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            else
                await page.EvaluateAsync(@"() => { const f = document.querySelector('form[action*=""bookpage_search""]'); if (f) f.submit(); }");
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        else if (isByPre1980)
        {
            var searchBtn = page.Locator("#searchPre").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            else
                await page.EvaluateAsync(@"() => { const f = document.querySelector('#searchFormPre, form[action*=""pre1980_search""]'); if (f) f.submit(); }");
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        else
        {
            var searchBtn = page.Locator("#searchD").First;
            if (await searchBtn.CountAsync() > 0)
                await searchBtn.ClickAsync(new LocatorClickOptions { Timeout = 15000 });
            await page.WaitForLoadStateAsync(LoadState.Load);
        }
        await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        Console.WriteLine($"[Search] Page URL after submit: {page.Url}");

        var resultErrorEl = page.GetByText("Record Count Exceeds").First;
        if (await resultErrorEl.CountAsync() > 0)
        {
            try
            {
                var errText = (await resultErrorEl.InnerTextAsync()).Trim();
                if (!string.IsNullOrEmpty(errText))
                {
                    Console.WriteLine($"[Result] {errText}");
                    throw new InvalidOperationException(errText);
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch { }
        }

        // ByDate, ByType, ByMunicipality, BySubdivision, BySTR: #dataTable2. ByName: table.table-select. ByBookPage/ByPre1980: #dataTable13. ByBookPage: ROW,TYPE,BOOK,PAGE,IMAGE(s). ByPre1980: ROW,TYPE,INSTRUMENT,IMAGE(s).
        var tableSelector = (isByBookPage || isByPre1980) ? "#dataTable13" : ((isByType || isByMunicipality || isBySubdivision || isBySTR) ? "#dataTable2" : (isByName ? "table.table-select" : "#dataTable2"));
        var tableTimeout = isFormSubmitMode ? 45000 : 20000;
        await page.WaitForSelectorAsync(tableSelector, new PageWaitForSelectorOptions { Timeout = tableTimeout });

        var grid = page.Locator(tableSelector);
        var headerCells = grid.Locator("thead th");
        var headerCount = await headerCells.CountAsync();
        var gridHeaders = new List<string>();
        for (var i = 0; i < headerCount; i++) gridHeaders.Add((await headerCells.Nth(i).InnerTextAsync()).Trim());
        Console.WriteLine($"[Grid] Headers: [{string.Join(", ", gridHeaders)}]");
        var fileNumberColIndex = gridHeaders.FindIndex(h => h.Equals("FILE NUMBER", StringComparison.OrdinalIgnoreCase));
        if (fileNumberColIndex < 0)
            for (var i = 0; i < gridHeaders.Count; i++)
                if (gridHeaders[i].ToUpperInvariant().Contains("FILE") && gridHeaders[i].ToUpperInvariant().Contains("NUMBER"))
                { fileNumberColIndex = i; break; }
        var fileDateColIndex = gridHeaders.FindIndex(h => h.Equals("FILE DATE", StringComparison.OrdinalIgnoreCase));
        var imageColIndex = gridHeaders.FindIndex(h => h.StartsWith("IMAGE", StringComparison.OrdinalIgnoreCase));
        if (imageColIndex < 0)
            imageColIndex = gridHeaders.FindIndex(h => { var u = h.ToUpperInvariant(); return u.Contains("IMAGE") || (u.Contains("VIEW") && !u.Contains("PREVIEW")); });
        var rowCount = await grid.Locator("tbody tr").CountAsync();
        Console.WriteLine($"[Grid] tableSelector={tableSelector}, fileNumberCol={fileNumberColIndex}, fileDateCol={fileDateColIndex}, imageCol={imageColIndex}, isFormSubmitMode={isFormSubmitMode}");

        return (context, page, rowCount, fileNumberColIndex, fileDateColIndex, imageColIndex, tableSelector, isFormSubmitMode);
    }
}

record CaptchaRect(double x, double y, double width, double height);
