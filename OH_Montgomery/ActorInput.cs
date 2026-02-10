using System.Text.Json.Serialization;

namespace OH_Montgomery;

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

    /// <summary>True when we should download images (ExportImagesFirst, ExportImagesOnly, ExportImageOnly, or All).</summary>
    public bool ShouldExportImages => string.Equals(ExportMode, "ExportImagesFirst", StringComparison.OrdinalIgnoreCase)
        || string.Equals(ExportMode, "ExportImagesOnly", StringComparison.OrdinalIgnoreCase)
        || string.Equals(ExportMode, "ExportImageOnly", StringComparison.OrdinalIgnoreCase)
        || string.Equals(ExportMode, "All", StringComparison.OrdinalIgnoreCase);
    public bool ExportImagesFirst => ShouldExportImages; // Legacy alias
    /// <summary>True when we should add image URLs to CSV (All only).</summary>
    public bool ExportAll => string.Equals(ExportMode, "All", StringComparison.OrdinalIgnoreCase);

    public string GetDateForFilename()
    {
        if (!string.IsNullOrWhiteSpace(StartDate) &&
            DateTime.TryParse(StartDate, out var d))
            return d.ToString("MM-dd-yyyy");
        return DateTime.Now.ToString("MM-dd-yyyy");
    }

    public void ValidateByName()
    {
        if (IsByName && string.IsNullOrWhiteSpace(LastName))
            throw new InvalidOperationException("LastName is required for 'By Name' search mode.");
    }

    public void ValidateByMunicipality()
    {
        if (!IsByMunicipality) return;
        var props = Properties ?? Municipality;
        if (string.IsNullOrWhiteSpace(props))
            throw new InvalidOperationException("Properties is required for 'By Municipality' search (e.g. DAYTON, BACHMAN).");
    }

    public void ValidateBySTR()
    {
        if (!IsBySTR) return;
        var hasSection = !string.IsNullOrWhiteSpace(Section);
        var hasTownship = !string.IsNullOrWhiteSpace(Township);
        var hasRange = !string.IsNullOrWhiteSpace(Range);
        if (!hasSection && !hasTownship && !hasRange)
            throw new InvalidOperationException("By STR requires at least one of: section, township, range (e.g. section: \"01\", township: \"1\", or range: \"5MRS\").");
    }

    public void ValidateBySubdivision()
    {
        if (!IsBySubdivision) return;
        var subs = Subdivisions ?? (Subdivision != null ? new[] { Subdivision } : null);
        if (subs == null || subs.Length == 0 || subs.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("Subdivisions (or subdivision) is required for 'By Subdivision' search (e.g. subdivisions: [\"ABBEY (LOTS 9 - 32)\"]).");
    }

    public void ValidateByInstrument()
    {
        if (!IsByInstrument) return;
        if (string.IsNullOrWhiteSpace(InstrumentYear) || InstrumentYear.Trim().Length != 4)
            throw new InvalidOperationException("InstrumentYear is required for 'By Instrument' search mode (4-digit year, e.g., 2024).");
        if (string.IsNullOrWhiteSpace(InstrumentNumber))
            throw new InvalidOperationException("InstrumentNumber is required for 'By Instrument' search mode.");
    }

    public void ValidateByBookPage()
    {
        if (IsByBookPage && (string.IsNullOrWhiteSpace(Book) || string.IsNullOrWhiteSpace(Page)))
            throw new InvalidOperationException("Book and Page are required for 'By Book/Page' search mode.");
    }

    public void ValidateByFiche()
    {
        if (IsByFiche && string.IsNullOrWhiteSpace(Fiche))
            throw new InvalidOperationException("Fiche is required for 'By Fiche' search mode.");
    }

    public void ValidateByPre1980()
    {
        if (IsByPre1980 && string.IsNullOrWhiteSpace(Pre1980Number))
            throw new InvalidOperationException("Pre1980Number is required for 'By Pre-1980' search mode.");
    }

    public void ValidateByType()
    {
        if (!IsByType) return;
        var types = DocumentTypes ?? (DocumentType != null ? new[] { DocumentType } : null);
        if (types == null || types.Length == 0 || types.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("At least one document type is required for 'By Type' search mode (e.g. documentTypes: [\"DEED\", \"MORTGAGE\"]).");
    }
}
