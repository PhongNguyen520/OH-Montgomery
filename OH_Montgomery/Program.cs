using Microsoft.Playwright;
using OH_Montgomery.Models;
using OH_Montgomery.Services;

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
            await ApifyHelper.SetStatusMessageAsync("Error: Failed to load input.", isTerminal: true);
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

        await ApifyHelper.SetStatusMessageAsync($"Starting scrape for {input.SearchMode}...");

        try
        {
            input.ValidateExportMode();
            input.ValidateByDate();
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
            await ApifyHelper.SetStatusMessageAsync($"Validation Error: {ex.Message}", isTerminal: true);
            Console.WriteLine($"ERROR: {ex.Message}");
            Environment.Exit(1);
        }

        var needsCaptcha = string.Equals(input.SearchMode, "ByDate", StringComparison.OrdinalIgnoreCase)
            || string.Equals(input.SearchMode, "ByInstrument", StringComparison.OrdinalIgnoreCase);
        if (needsCaptcha && string.IsNullOrWhiteSpace(input.TwoCaptchaApiKey) && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("TWO_CAPTCHA_API_KEY")))
        {
            await ApifyHelper.SetStatusMessageAsync("Error: Captcha required. Provide twoCaptchaApiKey.", isTerminal: true);
            Console.WriteLine("ERROR: By Date and By Instrument modes require captcha. Provide 'twoCaptchaApiKey' in input or set TWO_CAPTCHA_API_KEY env var.");
            Environment.Exit(1);
        }

        Console.WriteLine($"[1] Config loaded: SearchMode={input.SearchMode}, ExportMode={input.ExportMode}");
        Console.WriteLine();

        try
        {
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
            await new ScraperService().RunScrapeAsync(browser, input);
            }
            catch (Exception ex)
            {
            await ApifyHelper.SetStatusMessageAsync($"Fatal Error: {ex.Message}", isTerminal: true);
            Console.WriteLine($"ERROR: {ex}");
                throw;
            }
    }
}
