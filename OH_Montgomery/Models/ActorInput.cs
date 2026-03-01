using System.Text.Json.Serialization;

namespace OH_Montgomery.Models;

/// <summary>
/// Apify Actor input schema. Maps to input_schema.json.
/// </summary>
public class ActorInput
{
    // --- 1. General Configuration ---

    [JsonPropertyName("searchMode")]
    public string SearchMode { get; set; } = "ByDate";

    [JsonPropertyName("exportMode")]
    public string ExportMode { get; set; } = "ExportDataOnly";

    [JsonPropertyName("captchaCode")]
    public string? CaptchaCode { get; set; }

    /// <summary>2Captcha API key. If set, captcha is solved automatically instead of manual input.</summary>
    [JsonPropertyName("twoCaptchaApiKey")]
    public string? TwoCaptchaApiKey { get; set; }

    [JsonPropertyName("sortOrder")]
    public string SortOrder { get; set; } = "FILE DATE ASC";

    // --- 2. Date & Index Criteria ---

    [JsonPropertyName("startDate")]
    public string? StartDate { get; set; }

    [JsonPropertyName("endDate")]
    public string? EndDate { get; set; }

    [JsonPropertyName("indexTypes")]
    public string[] IndexTypes { get; set; } = ["Deeds", "Mortgages"];

    // --- 3. Mode Details: Name & Type ---

    [JsonPropertyName("lastName")]
    public string? LastName { get; set; }

    [JsonPropertyName("firstName")]
    public string? FirstName { get; set; }

    [JsonPropertyName("side")]
    public string Side { get; set; } = "Both";

    /// <summary>Last name search modifier: BeginWith, Contains, Exactly. Used by ByName and ByType.</summary>
    [JsonPropertyName("lastNameSearch")]
    public string? LastNameSearch { get; set; }

    /// <summary>First name search modifier: BeginWith, Contains, Exactly. Used by ByName and ByType.</summary>
    [JsonPropertyName("firstNameSearch")]
    public string? FirstNameSearch { get; set; }

    [JsonPropertyName("documentType")]
    public string? DocumentType { get; set; }

    /// <summary>ByType: array of document type names e.g. ["DEED", "MORTGAGE", "RELEASE OF MORTGAGE"]. At least one required.</summary>
    [JsonPropertyName("documentTypes")]
    public string[]? DocumentTypes { get; set; }

    [JsonPropertyName("includeFederalLien")]
    public bool IncludeFederalLien { get; set; } = false;

    // --- 4. Mode Details: Location ---

    [JsonPropertyName("municipality")]
    public string? Municipality { get; set; }

    /// <summary>By Municipality: the *PROPERTIES field on the form (e.g. DAYTON, BACHMAN).</summary>
    [JsonPropertyName("properties")]
    public string? Properties { get; set; }

    [JsonPropertyName("subdivision")]
    public string? Subdivision { get; set; }

    /// <summary>By Subdivision: array of *PROPERTIES (e.g. ["ABBEY (LOTS 9 - 32)", "ACCENT PARK S01"]). At least one required.</summary>
    [JsonPropertyName("subdivisions")]
    public string[]? Subdivisions { get; set; }

    [JsonPropertyName("lot")]
    public string? Lot { get; set; }

    [JsonPropertyName("section")]
    public string? Section { get; set; }

    [JsonPropertyName("township")]
    public string? Township { get; set; }

    [JsonPropertyName("range")]
    public string? Range { get; set; }

    // --- 5. Mode Details: Instrument & Fiche ---

    [JsonPropertyName("instrumentYear")]
    public string? InstrumentYear { get; set; }

    [JsonPropertyName("instrumentNumber")]
    public string? InstrumentNumber { get; set; }

    [JsonPropertyName("fiche")]
    public string? Fiche { get; set; }

    // --- 6. Mode Details: Book/Page & Pre-1980 ---

    [JsonPropertyName("book")]
    public string? Book { get; set; }

    [JsonPropertyName("page")]
    public string? Page { get; set; }

    [JsonPropertyName("pageThru")]
    public string? PageThru { get; set; }

    /// <summary>By Book/Page: Index Type. All (%), DEED/DBK, MTG/MBK, PLT/PBK.</summary>
    [JsonPropertyName("bookPageIndexType")]
    public string? BookPageIndexType { get; set; }

    [JsonPropertyName("pre1980Year")]
    public string? Pre1980Year { get; set; }

    [JsonPropertyName("pre1980Number")]
    public string? Pre1980Number { get; set; }

    /// <summary>By Pre-1980: Index Type. All (%), PDE (Deeds), PMT (Mortgages).</summary>
    [JsonPropertyName("pre1980IndexType")]
    public string? Pre1980IndexType { get; set; }

    // --- Helpers ---

    public bool IsByDate => string.Equals(SearchMode, "ByDate", StringComparison.OrdinalIgnoreCase);
    public bool IsByName => string.Equals(SearchMode, "ByName", StringComparison.OrdinalIgnoreCase);
    public bool IsByType => string.Equals(SearchMode, "ByType", StringComparison.OrdinalIgnoreCase);
    public bool IsByMunicipality => string.Equals(SearchMode, "ByMunicipality", StringComparison.OrdinalIgnoreCase);
    public bool IsBySubdivision => string.Equals(SearchMode, "BySubdivision", StringComparison.OrdinalIgnoreCase);
    public bool IsBySTR => string.Equals(SearchMode, "BySTR", StringComparison.OrdinalIgnoreCase);
    public bool IsByInstrument => string.Equals(SearchMode, "ByInstrument", StringComparison.OrdinalIgnoreCase);
    public bool IsByBookPage => string.Equals(SearchMode, "ByBookPage", StringComparison.OrdinalIgnoreCase);
    public bool IsByFiche => string.Equals(SearchMode, "ByFiche", StringComparison.OrdinalIgnoreCase);
    public bool IsByPre1980 => string.Equals(SearchMode, "ByPre1980", StringComparison.OrdinalIgnoreCase);

    /// <summary>True when we should download images (ExportImageOnly or All).</summary>
    public bool ShouldExportImages => string.Equals(ExportMode, "ExportImageOnly", StringComparison.OrdinalIgnoreCase)
        || string.Equals(ExportMode, "All", StringComparison.OrdinalIgnoreCase);
    public bool ExportImagesFirst => ShouldExportImages; // Legacy alias (kept for backward compatibility)
    /// <summary>True when we should add image URLs to CSV (All only).</summary>
    public bool ExportAll => string.Equals(ExportMode, "All", StringComparison.OrdinalIgnoreCase);

    public string GetDateForFilename()
    {
        if (!string.IsNullOrWhiteSpace(StartDate) &&
            DateTime.TryParse(StartDate, out var d))
            return d.ToString("MM-dd-yyyy");
        return DateTime.Now.ToString("MM-dd-yyyy");
    }

    public void ValidateExportMode()
    {
        if (string.IsNullOrWhiteSpace(ExportMode))
            throw new InvalidOperationException("exportMode is required.");

        var allowed = new[] { "ExportDataOnly", "ExportImageOnly", "All" };
        var match = allowed.FirstOrDefault(m => string.Equals(ExportMode, m, StringComparison.OrdinalIgnoreCase));
        if (match == null)
            throw new InvalidOperationException("exportMode must be one of: ExportDataOnly, ExportImageOnly, All.");

        // Normalize to canonical casing/value
        ExportMode = match;
    }

    public void ValidateByDate()
    {
        if (!IsByDate) return;

        if (string.IsNullOrWhiteSpace(StartDate))
            throw new InvalidOperationException("startDate is required for 'ByDate' search mode.");
        if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var start))
            throw new InvalidOperationException("startDate must use format yyyy-MM-dd (e.g. 2026-02-06).");

        if (string.IsNullOrWhiteSpace(EndDate))
            throw new InvalidOperationException("endDate is required for 'ByDate' search mode.");
        if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var end))
            throw new InvalidOperationException("endDate must use format yyyy-MM-dd (e.g. 2026-02-06).");

        if (start > end)
            throw new InvalidOperationException("startDate must be earlier than or equal to endDate for 'ByDate' search mode.");

        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'ByDate' search mode and must contain at least one index type.");

        // For ByDate, each IndexType must match one of the allowed values from the UI.
        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'ByDate' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        if (string.IsNullOrWhiteSpace(ExportMode))
            throw new InvalidOperationException("exportMode is required for 'ByDate' search mode.");

        // SortOrder for ByDate must match one of the Sort options in the UI.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var match = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException($"SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");

            // Normalize to canonical casing/value
            SortOrder = match;
        }
    }

    public void ValidateByName()
    {
        if (!IsByName) return;

        if (string.IsNullOrWhiteSpace(LastName))
            throw new InvalidOperationException("LastName is required for 'By Name' search mode.");

        // LastNameSearch / FirstNameSearch must match the UI options: BeginWith, Contains, Exactly.
        var allowedNameSearch = new[] { "BeginWith", "Contains", "Exactly" };
        if (!string.IsNullOrWhiteSpace(LastNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, LastNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("lastNameSearch must be one of: BeginWith, Contains, Exactly.");
            LastNameSearch = match;
        }
        if (!string.IsNullOrWhiteSpace(FirstNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, FirstNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("firstNameSearch must be one of: BeginWith, Contains, Exactly.");
            FirstNameSearch = match;
        }

        // Side must match the By Name UI: Both, One, Two.
        if (!string.IsNullOrWhiteSpace(Side))
        {
            var allowedSides = new[] { "Both", "One", "Two" };
            var match = allowedSides.FirstOrDefault(s => string.Equals(s, Side.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("side must be one of: Both, One, Two.");
            Side = match;
        }

        // IndexTypes for ByName: same allowed set as ByDate, minus Veteran Graves (not in UI here).
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Name' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Name' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // Optional date range: if provided, must respect yyyy-MM-dd and start <= end.
        if (!string.IsNullOrWhiteSpace(StartDate))
        {
            if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Name' search mode, startDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(EndDate))
        {
            if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Name' search mode, endDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(StartDate) && !string.IsNullOrWhiteSpace(EndDate)
            && DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var s)
            && DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var e)
            && s > e)
        {
            throw new InvalidOperationException("For 'By Name' search mode, startDate must be earlier than or equal to endDate when both are provided.");
        }

        // SortOrder: must match one of the UI options.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var match = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException($"For 'By Name' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = match;
        }
    }

    public void ValidateByMunicipality()
    {
        if (!IsByMunicipality) return;
        var props = Properties ?? Municipality;
        if (string.IsNullOrWhiteSpace(props))
            throw new InvalidOperationException("Properties is required for 'By Municipality' search (e.g. DAYTON, BACHMAN).");

        // Properties must match one of the options in the Municipality UI.
        var allowedProperties = new[]
        {
            "AMITYVILLE",
            "ARLINGTON",
            "BACHMAN",
            "BEAVERTOWN",
            "BROOKVILLE",
            "CARLISLE",
            "CENTERVILLE",
            "CHAMBERSBURG",
            "CLAY TOWNSHIP",
            "CLAYTON",
            "DAYTON",
            "DODSON",
            "ENGLEWOOD",
            "FARMERSVILLE",
            "GERMAN TOWNSHIP",
            "GERMANTOWN",
            "HARRISON TOWNSHIP",
            "JEFFERSON TOWNSHIP",
            "JOHNSVILLE",
            "KETTERING",
            "LIBERTY",
            "LITTLE YORK",
            "MADISON TOWNSHIP",
            "MIAMISBURG",
            "MORAINE",
            "MURLIN HEIGHTS",
            "NEW LEBANON",
            "OAKWOOD",
            "PHILLIPSBURG",
            "PYRMONT",
            "RIVERSIDE",
            "SALEM",
            "SPRINGBORO (MONTGOMERY & WARREN)",
            "SUNBURY",
            "TROTWOOD",
            "UNION",
            "VANDALIA",
            "VERONA",
            "WEST CARROLLTON",
            "WOODBOURNE"
        };
        var propMatch = allowedProperties.FirstOrDefault(p => string.Equals(p, props.Trim(), StringComparison.OrdinalIgnoreCase));
        if (propMatch == null)
            throw new InvalidOperationException($"Properties value '{props}' is invalid for 'By Municipality' search.");
        Properties = propMatch;

        // Index Type*: same allowed set as ByDate UI.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Municipality' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Municipality' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // Lot: optional, UI already constrains format; we only enforce length via HTML, so no extra validation here.

        // SortOrder: same options as other modes using Sort.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var sortMatch = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (sortMatch == null)
                throw new InvalidOperationException($"For 'By Municipality' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = sortMatch;
        }

        // Optional extra filters: Last/First name search and date range behave like ByName.
        var allowedNameSearch = new[] { "BeginWith", "Contains", "Exactly" };
        if (!string.IsNullOrWhiteSpace(LastNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, LastNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Municipality' search mode, lastNameSearch must be one of: BeginWith, Contains, Exactly.");
            LastNameSearch = match;
        }
        if (!string.IsNullOrWhiteSpace(FirstNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, FirstNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Municipality' search mode, firstNameSearch must be one of: BeginWith, Contains, Exactly.");
            FirstNameSearch = match;
        }

        if (!string.IsNullOrWhiteSpace(StartDate))
        {
            if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Municipality' search mode, startDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(EndDate))
        {
            if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Municipality' search mode, endDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(StartDate) && !string.IsNullOrWhiteSpace(EndDate)
            && DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var s)
            && DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var e)
            && s > e)
        {
            throw new InvalidOperationException("For 'By Municipality' search mode, startDate must be earlier than or equal to endDate when both are provided.");
        }
    }

    public void ValidateBySTR()
    {
        if (!IsBySTR) return;
        var hasSection = !string.IsNullOrWhiteSpace(Section);
        var hasTownship = !string.IsNullOrWhiteSpace(Township);
        var hasRange = !string.IsNullOrWhiteSpace(Range);
        if (!hasSection && !hasTownship && !hasRange)
            throw new InvalidOperationException("By STR requires at least one of: section, township, range (e.g. section: \"01\", township: \"1\", or range: \"5MRS\").");

        // Properties*: must be one of the STR UI options (including ALL).
        if (string.IsNullOrWhiteSpace(Properties) && string.IsNullOrWhiteSpace(Municipality))
            throw new InvalidOperationException("Properties is required for 'By STR' search mode (e.g. DAYTON, MIAMI TOWNSHIP, ALL).");

        var props = Properties ?? Municipality;
        var allowedPropertiesForStr = new[]
        {
            "ALL",
            "BACHMAN",
            "BROOKVILLE",
            "BUTLER TOWNSHIP",
            "CARLISLE",
            "CENTERVILLE",
            "CHAMBERSBURG",
            "CLAY TOWNSHIP",
            "CLAYTON",
            "DAYTON",
            "ENGLEWOOD",
            "FARMERSVILLE",
            "GERMAN TOWNSHIP",
            "GERMANTOWN",
            "HARRISON TOWNSHIP",
            "HUBER HEIGHTS",
            "JACKSON TOWNSHIP",
            "JEFFERSON TOWNSHIP",
            "JOHNSVILLE",
            "KETTERING",
            "LITTLE YORK",
            "MAD RIVER TOWNSHIP",
            "MADISON TOWNSHIP",
            "MIAMI TOWNSHIP",
            "MIAMISBURG",
            "MORAINE",
            "MURLIN HEIGHTS",
            "NEW LEBANON",
            "OAKWOOD",
            "PERRY TOWNSHIP",
            "PHILLIPSBURG",
            "PYRMONT",
            "RANDOLPH TOWNSHIP",
            "RIVERSIDE",
            "SPRINGBORO (MONTGOMERY & WARREN)",
            "SUNBURY",
            "TROTWOOD",
            "UNION",
            "VANDALIA",
            "VERONA",
            "WASHINGTON TOWNSHIP",
            "WEST CARROLLTON",
            "WOODBOURNE"
        };
        var propMatch = allowedPropertiesForStr.FirstOrDefault(p => string.Equals(p, props!.Trim(), StringComparison.OrdinalIgnoreCase));
        if (propMatch == null)
            throw new InvalidOperationException($"Properties value '{props}' is invalid for 'By STR' search.");
        Properties = propMatch;

        // Index Type*: same allowed set as ByDate UI.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By STR' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By STR' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // SortOrder: same options as other modes using Sort.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var sortMatch = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (sortMatch == null)
                throw new InvalidOperationException($"For 'By STR' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = sortMatch;
        }

        // Optional extra filters: Last/First name search and date range behave like ByName.
        var allowedNameSearch = new[] { "BeginWith", "Contains", "Exactly" };
        if (!string.IsNullOrWhiteSpace(LastNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, LastNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By STR' search mode, lastNameSearch must be one of: BeginWith, Contains, Exactly.");
            LastNameSearch = match;
        }
        if (!string.IsNullOrWhiteSpace(FirstNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, FirstNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By STR' search mode, firstNameSearch must be one of: BeginWith, Contains, Exactly.");
            FirstNameSearch = match;
        }

        if (!string.IsNullOrWhiteSpace(StartDate))
        {
            if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By STR' search mode, startDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(EndDate))
        {
            if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By STR' search mode, endDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(StartDate) && !string.IsNullOrWhiteSpace(EndDate)
            && DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var s2)
            && DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var e2)
            && s2 > e2)
        {
            throw new InvalidOperationException("For 'By STR' search mode, startDate must be earlier than or equal to endDate when both are provided.");
        }
    }

    public void ValidateBySubdivision()
    {
        if (!IsBySubdivision) return;
        var subs = Subdivisions ?? (Subdivision != null ? new[] { Subdivision } : null);
        if (subs == null || subs.Length == 0 || subs.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("Subdivisions (or subdivision) is required for 'By Subdivision' search (e.g. subdivisions: [\"ABBEY (LOTS 9 - 32)\"]).");

        // We do not validate subdivision values against a fixed keyword list; UI controls this via modal.

        // Index Type*: same allowed set as ByDate UI.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Subdivision' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Subdivision' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // SortOrder: same options as other modes using Sort.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var sortMatch = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (sortMatch == null)
                throw new InvalidOperationException($"For 'By Subdivision' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = sortMatch;
        }

        // Optional extra filters: Last/First name search and date range behave like ByName.
        var allowedNameSearch = new[] { "BeginWith", "Contains", "Exactly" };
        if (!string.IsNullOrWhiteSpace(LastNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, LastNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Subdivision' search mode, lastNameSearch must be one of: BeginWith, Contains, Exactly.");
            LastNameSearch = match;
        }
        if (!string.IsNullOrWhiteSpace(FirstNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, FirstNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Subdivision' search mode, firstNameSearch must be one of: BeginWith, Contains, Exactly.");
            FirstNameSearch = match;
        }

        if (!string.IsNullOrWhiteSpace(StartDate))
        {
            if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Subdivision' search mode, startDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(EndDate))
        {
            if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Subdivision' search mode, endDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(StartDate) && !string.IsNullOrWhiteSpace(EndDate)
            && DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var s)
            && DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var e)
            && s > e)
        {
            throw new InvalidOperationException("For 'By Subdivision' search mode, startDate must be earlier than or equal to endDate when both are provided.");
        }
    }

    public void ValidateByInstrument()
    {
        if (!IsByInstrument) return;
        if (string.IsNullOrWhiteSpace(InstrumentYear) || InstrumentYear.Trim().Length != 4)
            throw new InvalidOperationException("InstrumentYear is required for 'By Instrument' search mode (4-digit year, e.g., 2024).");
        if (string.IsNullOrWhiteSpace(InstrumentNumber))
            throw new InvalidOperationException("InstrumentNumber is required for 'By Instrument' search mode.");

        // Index Type*: cannot be empty when using ByInstrument.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Instrument' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Instrument' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // SortOrder: same options as other modes using Sort.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var sortMatch = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (sortMatch == null)
                throw new InvalidOperationException($"For 'By Instrument' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = sortMatch;
        }
    }

    public void ValidateByBookPage()
    {
        if (!IsByBookPage) return;

        // Book*: required, letters/numbers only (matches UI pattern).
        if (string.IsNullOrWhiteSpace(Book))
            throw new InvalidOperationException("Book is required for 'By Book/Page' search mode.");
        if (!Book.All(c => char.IsLetterOrDigit(c)))
            throw new InvalidOperationException("Book must contain letters and/or numbers only for 'By Book/Page' search mode.");

        // Page*: required, numeric only.
        if (string.IsNullOrWhiteSpace(Page))
            throw new InvalidOperationException("Page is required for 'By Book/Page' search mode.");
        if (!Page.All(char.IsDigit))
            throw new InvalidOperationException("Page must be numeric for 'By Book/Page' search mode.");

        // Thru: optional, but if present must be numeric.
        if (!string.IsNullOrWhiteSpace(PageThru) && !PageThru.All(char.IsDigit))
            throw new InvalidOperationException("PageThru must be numeric when provided for 'By Book/Page' search mode.");

        // IndexType (BookPageIndexType): optional, but if provided must map to IndexTypeToRissBookPage keys.
        if (!string.IsNullOrWhiteSpace(BookPageIndexType))
        {
            var allowed = new[]
            {
                "All", "%",
                "DEED", "DBK", "Deeds",
                "MTG", "MBK", "Mortgages", "Mortgage",
                "PLT", "PBK"
            };
            var match = allowed.FirstOrDefault(v => string.Equals(v, BookPageIndexType.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("bookPageIndexType must be one of: All, DEED, MTG, PLT (or their DBK/MBK/PBK/% equivalents).");
            BookPageIndexType = match;
        }
    }

    public void ValidateByFiche()
    {
        if (!IsByFiche) return;

        // Fiche*: required, 6–10 alphanumeric characters.
        if (string.IsNullOrWhiteSpace(Fiche))
            throw new InvalidOperationException("Fiche is required for 'By Fiche' search mode.");
        var ficheTrim = Fiche.Trim();
        if (ficheTrim.Length < 6 || ficheTrim.Length > 10 || !ficheTrim.All(char.IsLetterOrDigit))
            throw new InvalidOperationException("Fiche must be between 6 and 10 alphanumeric characters for 'By Fiche' search mode (e.g. 910022a05).");
        Fiche = ficheTrim.ToUpperInvariant();

        // Index Type*: Deeds, Mortgages, UCC.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Fiche' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "UCC"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Fiche' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // SortOrder: FILE DATE ASC or FILE DATE DESC.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "FILE DATE ASC",
                "FILE DATE DESC"
            };
            var match = allowedSortOrders.FirstOrDefault(s => string.Equals(s, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException($"For 'By Fiche' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = match;
        }
    }

    public void ValidateByPre1980()
    {
        if (!IsByPre1980) return;

        // Pre 1980 Number*: required, numeric only.
        if (string.IsNullOrWhiteSpace(Pre1980Number))
            throw new InvalidOperationException("Pre1980Number is required for 'By Pre-1980' search mode.");
        var num = Pre1980Number.Trim();
        if (!num.All(char.IsDigit))
            throw new InvalidOperationException("Pre1980Number must be numeric for 'By Pre-1980' search mode.");
        Pre1980Number = num;

        // Pre 1980 Year: optional in UI; if provided, must be between 1971 and 1979.
        if (!string.IsNullOrWhiteSpace(Pre1980Year))
        {
            var yearTrim = Pre1980Year.Trim();
            if (!int.TryParse(yearTrim, out var y) || y < 1971 || y > 1979)
                throw new InvalidOperationException("Pre1980Year (if provided) must be a 4-digit year between 1971 and 1979 for 'By Pre-1980' search mode.");
            Pre1980Year = yearTrim;
        }

        // Index Type: All (default), PDE (Deeds), PMT (Mortgages).
        if (!string.IsNullOrWhiteSpace(Pre1980IndexType))
        {
            var allowed = new[] { "All", "%", "PDE", "PMT" };
            var match = allowed.FirstOrDefault(v => string.Equals(v, Pre1980IndexType.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("pre1980IndexType must be one of: All, PDE, PMT.");
            Pre1980IndexType = match;
        }
    }

    public void ValidateByType()
    {
        if (!IsByType) return;

        // Document Type*: required but free-form (must map to RISS options later; we do not validate keywords here).
        var types = DocumentTypes ?? (DocumentType != null ? new[] { DocumentType } : null);
        if (types == null || types.Length == 0 || types.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("At least one document type is required for 'By Type' search mode (e.g. documentTypes: [\"DEED\", \"MORTGAGE\"]).");

        // Index Type*: same allowed set as ByDate UI.
        if (IndexTypes == null || IndexTypes.Length == 0 || IndexTypes.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("indexTypes is required for 'By Type' search mode and must contain at least one index type.");

        var allowedIndexTypes = new[]
        {
            "Deeds",
            "Mortgages",
            "Partnerships",
            "UCC",
            "Service Discharge",
            "Veteran Graves"
        };
        var invalidIndex = IndexTypes
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .FirstOrDefault(t => !allowedIndexTypes.Any(a => string.Equals(a, t.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (invalidIndex != null)
            throw new InvalidOperationException($"Invalid indexTypes value for 'By Type' search mode: '{invalidIndex}'. Allowed values: {string.Join(", ", allowedIndexTypes)}.");

        // LastNameSearch / FirstNameSearch: same 3 options as ByName.
        var allowedNameSearch = new[] { "BeginWith", "Contains", "Exactly" };
        if (!string.IsNullOrWhiteSpace(LastNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, LastNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Type' search mode, lastNameSearch must be one of: BeginWith, Contains, Exactly.");
            LastNameSearch = match;
        }
        if (!string.IsNullOrWhiteSpace(FirstNameSearch))
        {
            var match = allowedNameSearch.FirstOrDefault(s => string.Equals(s, FirstNameSearch.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Type' search mode, firstNameSearch must be one of: BeginWith, Contains, Exactly.");
            FirstNameSearch = match;
        }

        // Side(s): Both / One / Two.
        if (!string.IsNullOrWhiteSpace(Side))
        {
            var allowedSides = new[] { "Both", "One", "Two" };
            var match = allowedSides.FirstOrDefault(s => string.Equals(s, Side.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException("For 'By Type' search mode, side must be one of: Both, One, Two.");
            Side = match;
        }

        // Optional date range: yyyy-MM-dd and start <= end when both provided.
        if (!string.IsNullOrWhiteSpace(StartDate))
        {
            if (!DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Type' search mode, startDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(EndDate))
        {
            if (!DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out _))
                throw new InvalidOperationException("For 'By Type' search mode, endDate (if provided) must use format yyyy-MM-dd (e.g. 2026-02-06).");
        }
        if (!string.IsNullOrWhiteSpace(StartDate) && !string.IsNullOrWhiteSpace(EndDate)
            && DateTime.TryParseExact(StartDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var s)
            && DateTime.TryParseExact(EndDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var e)
            && s > e)
        {
            throw new InvalidOperationException("For 'By Type' search mode, startDate must be earlier than or equal to endDate when both are provided.");
        }

        // SortOrder: same options as other modes using Sort.
        if (!string.IsNullOrWhiteSpace(SortOrder))
        {
            var allowedSortOrders = new[]
            {
                "MATCH ASC",
                "FILE DATE ASC",
                "FILE NUMBER ASC",
                "TYPE ASC",
                "MATCH DESC",
                "FILE DATE DESC",
                "FILE NUMBER DESC",
                "TYPE DESC"
            };
            var match = allowedSortOrders.FirstOrDefault(x => string.Equals(x, SortOrder.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match == null)
                throw new InvalidOperationException($"For 'By Type' search mode, SortOrder must be one of: {string.Join(", ", allowedSortOrders)}.");
            SortOrder = match;
        }
    }
}
