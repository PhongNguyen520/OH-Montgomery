using System.Globalization;
using System.Text.RegularExpressions;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Playwright;
using OH_Montgomery.Models;
using OH_Montgomery.Utils;

namespace OH_Montgomery.Services;

/// <summary>Core scraping logic: context creation, search, and record extraction.</summary>
public class ScraperService
{
    public async Task RunScrapeAsync(IBrowser browser, ActorInput input)
    {
        var needsCaptcha = string.Equals(input.SearchMode, "ByDate", StringComparison.OrdinalIgnoreCase)
            || string.Equals(input.SearchMode, "ByInstrument", StringComparison.OrdinalIgnoreCase);

        IBrowserContext? context = null;
        IPage? page = null;
        int rowCount = 0, fileNumberColIndex = 0, fileDateColIndex = 0, imageColIndex = 0;
        string tableSelector = "";
        bool isFormSubmitMode = false;

        int searchRetries = 3;
        bool searchSuccess = false;

        for (int attempt = 1; attempt <= searchRetries; attempt++)
        {
            try
            {
                await ApifyHelper.SetStatusMessageAsync($"Search attempt {attempt} of {searchRetries}...");
                (context, page, rowCount, fileNumberColIndex, fileDateColIndex, imageColIndex, tableSelector, isFormSubmitMode) =
                    await CreateContextAndLoadResultsAsync(browser, input, needsCaptcha);
                searchSuccess = true;
                break;
            }
            catch (Exception ex)
            {
                var is500Limit = ex is InvalidOperationException && ex.Message.Contains("Record Count Exceeds", StringComparison.OrdinalIgnoreCase);
                if (is500Limit)
                {
                    await ApifyHelper.SetStatusMessageAsync($"Error: {ex.Message}", isTerminal: true);
                    Console.WriteLine($"    Error: {ex.Message} - Exiting.");
                    await browser.CloseAsync();
                    return;
                }

                Console.WriteLine($"    [Attempt {attempt}] Search failed: {ex.Message}");
                if (context != null) { try { await context.CloseAsync(); } catch { } }

                if (attempt == searchRetries)
                {
                    await ApifyHelper.SetStatusMessageAsync($"Fatal Error during search after {searchRetries} attempts: {ex.Message}", isTerminal: true);
                    await browser.CloseAsync();
                    return;
                }

                await Task.Delay(5000);
            }
        }

        if (!searchSuccess) return;
        Console.WriteLine("    Grid loaded.");
        Console.WriteLine();

        var grid = page!.Locator(tableSelector);
        if (await grid.CountAsync() > 0)
        {
            var dataRows = grid.Locator("tbody tr");
            Console.WriteLine($"    Found {rowCount} rows.");

            if (rowCount == 0)
            {
                await ApifyHelper.SetStatusMessageAsync("Finished: No records found for the given criteria.", isTerminal: true);
                Console.WriteLine("    No records. Exiting.");
                await browser.CloseAsync();
                return;
            }

            const int TestLimit = 5;
            var rowsToProcess = Math.Min(rowCount, TestLimit);
            if (rowCount > TestLimit)
                Console.WriteLine($"[OH_Montgomery] Limiting to {TestLimit} records for test (total {rowCount} skipped).");

            await ApifyHelper.SetStatusMessageAsync($"Processing {rowsToProcess} records...");

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
                try { await context!.GrantPermissionsAsync(new[] { "clipboard-read", "clipboard-write", "storage-access" }, new BrowserContextGrantPermissionsOptions { Origin = "https://onbase.mcohio.org" }); } catch { }

            Console.WriteLine($"    ExportMode={input.ExportMode}, imageColIndex={imageColIndex}, shouldDownloadImages={shouldDownloadImages}");
            Console.WriteLine($"    Processing rows: image + detail per row, batch 10 to CSV (Images/{dateForFilename}/...)");
            var buffer = new List<Record>();
            var kvStore = Path.Combine(baseDir, "apify_storage", "key_value_store");
            Directory.CreateDirectory(kvStore);
            var outputPath = Path.Combine(kvStore, $"OH_Montgomery_{dateForFilename}.csv");
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = "|", HasHeaderRecord = true, ShouldQuote = args => true };
            var succeeded = 0;
            var failed = 0;

            await using (var writer = new StreamWriter(outputPath, append: false))
            using (var csv = new CsvWriter(writer, csvConfig))
            {
                for (var r = 0; r < rowsToProcess; r++)
                {
                    if ((r + 1) % 10 == 0)
                    {
                        await ApifyHelper.SetStatusMessageAsync($"Processing row {r + 1} of {rowsToProcess}...");
                        Console.WriteLine($"    Row {r + 1}/{rowsToProcess}...");
                    }
                    grid = page!.Locator(tableSelector);
                    dataRows = grid.Locator("tbody tr");
                    var row = dataRows.Nth(r);
                    var cells = row.Locator("td");
                    var cellCount = await cells.CountAsync();
                    if (cellCount == 0) continue;

                    try
                    {
                    var record = new Record();
                    record.RecordingDate = fileDateColIndex >= 0 && fileDateColIndex < cellCount ? (await cells.Nth(fileDateColIndex).InnerTextAsync()).Trim() : "";

                    if (input.IsByBookPage && cellCount >= 4)
                    {
                        record.DocumentType = (await cells.Nth(1).InnerTextAsync()).Trim();
                        record.Book = (await cells.Nth(2).InnerTextAsync()).Trim();
                        record.Page = (await cells.Nth(3).InnerTextAsync()).Trim();
                        record.BookType = record.DocumentType;
                        record.DocumentNumber = $"{record.Book}-{record.Page}".TrimEnd('-');
                    }
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
                                imgPage = await context!.RunAndWaitForPageAsync(async () => { await cameraLink.ClickAsync(); }, new BrowserContextRunAndWaitForPageOptions { Timeout = 30000 });
                                await imgPage.BringToFrontAsync();
                                imgPage.SetDefaultTimeout(35000);
                                await imgPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                                await imgPage.WaitForTimeoutAsync(2000);
                                await imgPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                                await imgPage.WaitForTimeoutAsync(6000);

                                try
                                {
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
                                        var folderRows = selectListFrame.Locator("#primaryHitlist_grid tbody tr");
                                        var folderCount = await folderRows.CountAsync();

                                        for (var frIndex = 0; frIndex < folderCount; frIndex++)
                                        {
                                            var rowFolder = folderRows.Nth(frIndex);
                                            var firstCell = rowFolder.Locator("td").First;

                                            if (await firstCell.CountAsync() == 0)
                                                continue;

                                            try
                                            {
                                                await firstCell.DblClickAsync();
                                            }
                                            catch
                                            {
                                                try
                                                {
                                                    await firstCell.ClickAsync();
                                                    await imgPage.WaitForTimeoutAsync(200);
                                                    await firstCell.ClickAsync();
                                                }
                                                catch { }
                                            }

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
                                                try
                                                {
                                                    var bytes = ImageProcessor.CreateMultiPageCompressedTiff(base64ArrFolder);
                                                    var fileName = $"{baseNameForFolder}.tif";
                                                    var kvKey = $"Images/{dateForFilename}/{baseNameForFolder}/{fileName}";

                                                    await ApifyHelper.SaveImageAsync(kvKey, bytes);
                                                    if (input.ExportAll) imageLinks.Add(ApifyHelper.GetRecordUrl(kvKey));

                                                    Console.WriteLine($"      Saved Multi-page TIF ({base64ArrFolder.Length} pages): {kvKey}");
                                                }
                                                catch (Exception ex) { Console.WriteLine($"      Save error: {ex.Message}"); }
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
                                    try
                                    {
                                        var bytes = ImageProcessor.CreateMultiPageCompressedTiff(base64Arr);
                                        var fileName = $"{baseName}.tif";
                                        var kvKey = $"Images/{dateForFilename}/{baseName}/{fileName}";

                                        await ApifyHelper.SaveImageAsync(kvKey, bytes);
                                        if (input.ExportAll) imageLinks.Add(ApifyHelper.GetRecordUrl(kvKey));

                                        Console.WriteLine($"      Saved Multi-page TIF ({base64Arr.Length} pages): {kvKey}");
                                    }
                                    catch (Exception ex) { Console.WriteLine($"      Save error: {ex.Message}"); }
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
                                    var navTask = page.WaitForURLAsync(url => url.Contains("docinfo", StringComparison.OrdinalIgnoreCase), new PageWaitForURLOptions { Timeout = 30000 });
                                    await clickTarget.ClickAsync();
                                    await navTask;
                                    await page.WaitForSelectorAsync("span.informationTitle", new PageWaitForSelectorOptions { Timeout = 15000 });
                                    record.DocumentNumber = await DomHelper.GetDetailValue(page, "INSTRUMENT");
                                    record.BookType = await DomHelper.GetDetailValue(page, "INDEX");
                                    record.DocumentType = await DomHelper.GetDetailValue(page, "TYPE");
                                    record.Amount = await DomHelper.GetDetailValue(page, "AMOUNT");
                                    var grantor = await DomHelper.GetDetailValueList(page, "GRANTORS");
                                    if (string.IsNullOrEmpty(grantor)) grantor = await DomHelper.GetDetailValueList(page, "MORTGAGORS");
                                    record.Grantor = grantor;
                                    record.Grantee = await DomHelper.GetGranteeValueList(page);
                                    record.Reference = await DomHelper.GetDetailValueList(page, "REFERENCES");
                                    record.Remarks = await DomHelper.GetRemarksList(page);
                                    record.Legal = await DomHelper.GetDetailValueList(page, "LEGAL");
                                    record.Property = await DomHelper.GetDetailValueList(page, "PROPERTY");
                                    await page.GoBackAsync(new PageGoBackOptions { WaitUntil = WaitUntilState.NetworkIdle });
                                    await page.WaitForSelectorAsync(tableSelector, new PageWaitForSelectorOptions { Timeout = 15000 });
                                    await page.WaitForTimeoutAsync(500);
                                }
                                else
                                {
                                    var detailPage = await context!.RunAndWaitForPageAsync(async () =>
                                    {
                                        await clickTarget.ClickAsync(new LocatorClickOptions { Modifiers = new[] { KeyboardModifier.Control } });
                                    }, new BrowserContextRunAndWaitForPageOptions { Timeout = 30000 });
                                    try
                                    {
                                        await detailPage.WaitForSelectorAsync("span.informationTitle", new PageWaitForSelectorOptions { Timeout = 15000 });
                                        record.DocumentNumber = await DomHelper.GetDetailValue(detailPage, "INSTRUMENT");
                                        record.BookType = await DomHelper.GetDetailValue(detailPage, "INDEX");
                                        record.DocumentType = await DomHelper.GetDetailValue(detailPage, "TYPE");
                                        record.Amount = await DomHelper.GetDetailValue(detailPage, "AMOUNT");
                                        var grantor = await DomHelper.GetDetailValueList(detailPage, "GRANTORS");
                                        if (string.IsNullOrEmpty(grantor)) grantor = await DomHelper.GetDetailValueList(detailPage, "MORTGAGORS");
                                        record.Grantor = grantor;
                                        record.Grantee = await DomHelper.GetGranteeValueList(detailPage);
                                        record.Reference = await DomHelper.GetDetailValueList(detailPage, "REFERENCES");
                                        record.Remarks = await DomHelper.GetRemarksList(detailPage);
                                        record.Legal = await DomHelper.GetDetailValueList(detailPage, "LEGAL");
                                        record.Property = await DomHelper.GetDetailValueList(detailPage, "PROPERTY");
                                    }
                                    finally { await detailPage.CloseAsync(); }
                                }
                            }
                            catch { }
                        }
                    }
                    buffer.Add(record);
                    succeeded++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    Row {r + 1} failed: {ex.Message}");
                        failed++;
                        continue;
                    }

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

            await ApifyHelper.SetStatusMessageAsync($"Finished! Total {rowsToProcess} requests: {succeeded} succeeded, {failed} failed.", isTerminal: true);
            Console.WriteLine($"    Exported {rowCount} records to CSV and Apify dataset.");
            Console.WriteLine($"    Also saved CSV: {outputPath}");
        }
        else
        {
            await ApifyHelper.SetStatusMessageAsync("Finished: No table or grid found.", isTerminal: true);
            Console.WriteLine($"    Table ({tableSelector}) not found or empty. Exiting.");
            await browser.CloseAsync();
            return;
        }

        Console.WriteLine();
        Console.WriteLine("[12] Done. Closing browser...");
        await browser.CloseAsync();
    }

    static async Task<(IBrowserContext ctx, IPage page, int rowCount, int fileNumberColIndex, int fileDateColIndex, int imageColIndex, string tableSelector, bool isByName)> CreateContextAndLoadResultsAsync(
        IBrowser browser, ActorInput input, bool needsCaptcha, bool hasRetriedCaptcha = false)
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
            await page.WaitForTimeoutAsync(3000);
        }

        await FormFiller.ClickSidebarTabAsync(page, input.SearchMode);
        await page.WaitForTimeoutAsync(1500);

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
            var captchaBase64 = await DomHelper.ExtractCaptchaBase64Async(page);
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
            case "ByDate": await FormFiller.FillByDateFormBatchAsync(page, input, needsCaptcha); break;
            case "ByName": await FormFiller.FillByNameFormBatchAsync(page, input); break;
            case "ByType": await FormFiller.FillByTypeForm(page, input); break;
            case "ByMunicipality": await FormFiller.FillByMunicipalityForm(page, input); break;
            case "BySubdivision": await FormFiller.FillBySubdivisionForm(page, input); break;
            case "BySTR": await FormFiller.FillBySTRForm(page, input); break;
            case "ByInstrument":
                await FormFiller.FillByInstrumentForm(page, input);
                if (needsCaptcha && !string.IsNullOrWhiteSpace(input.CaptchaCode)) await FormFiller.FillCaptchaAsync(page, input.CaptchaCode);
                break;
            case "ByBookPage": await FormFiller.FillByBookPageForm(page, input); break;
            case "ByFiche": await FormFiller.FillByFicheForm(page, input); break;
            case "ByPre1980": await FormFiller.FillByPre1980Form(page, input); break;
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

        if (needsCaptcha)
        {
            var badCaptcha = page.Locator("#badCaptcha").First;
            if (await badCaptcha.CountAsync() > 0 && await badCaptcha.IsVisibleAsync())
            {
                Console.WriteLine("[Captcha] Incorrect captcha detected after submit.");
                if (!hasRetriedCaptcha)
                {
                    Console.WriteLine("[Captcha] Reloading and retrying search one more time...");
                    try { await context.CloseAsync(); } catch { }
                    return await CreateContextAndLoadResultsAsync(browser, input, needsCaptcha, true);
                }

                throw new InvalidOperationException("Captcha was incorrect even after retrying once.");
            }
        }

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
