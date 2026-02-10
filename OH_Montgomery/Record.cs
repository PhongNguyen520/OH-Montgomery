using CsvHelper.Configuration.Attributes;

namespace OH_Montgomery;

/// <summary>
/// Scraped record per Data File Formatting Rules. 15 columns, pipe delimiter, semicolon for nested lists.
/// </summary>
public class Record
{
    [Name("Document Number")]
    public string DocumentNumber { get; set; } = "";   // Map from Instrument

    [Name("Book")]
    public string Book { get; set; } = "";             // Always empty

    [Name("Page")]
    public string Page { get; set; } = "";             // Always empty

    [Name("Recording Date")]
    public string RecordingDate { get; set; } = "";    // Map from File Date (grid)

    [Name("Book Type")]
    public string BookType { get; set; } = "";         // Map from Index

    [Name("Document Type")]
    public string DocumentType { get; set; } = "";     // Map from Type

    [Name("Amount")]
    public string Amount { get; set; } = "";

    [Name("Grantor")]
    public string Grantor { get; set; } = "";          // Map from Mortgagors; join multiple with ;

    [Name("Grantee")]
    public string Grantee { get; set; } = "";          // Map from Mortgagees; join multiple with ;

    [Name("Reference")]
    public string Reference { get; set; } = "";        // Map from References; join multiple with ;

    [Name("Remarks")]
    public string Remarks { get; set; } = "";          // Map from Note; join with ;

    [Name("Parcel Number")]
    public string ParcelNumber { get; set; } = "";     // Always empty

    [Name("Legal")]
    public string Legal { get; set; } = "";            // Join multiple lines with ;

    [Name("Property Address")]
    public string PropertyAddress { get; set; } = "";  // Always empty

    [Name("Property")]
    public string Property { get; set; } = "";         // Join multiple lines with ;

    [Name("Image Links")]
    public string ImageLinks { get; set; } = "";       // Newline-separated URLs; CSV-quoted when written
}
