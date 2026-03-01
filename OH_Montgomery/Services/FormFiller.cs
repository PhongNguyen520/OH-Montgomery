using System.Text.Json;
using Microsoft.Playwright;
using OH_Montgomery.Models;

namespace OH_Montgomery.Services;

/// <summary>Form filling utilities for OH Montgomery RISS search forms.</summary>
public static class FormFiller
{
    /// <summary>RISS By Name: Side values %=Both, 1=One, 2=Two. Web UI uses Both/Grantor/Grantee.</summary>
    public static readonly Dictionary<string, string> SideToRissByName = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Both"] = "%", ["Grantor"] = "1", ["Grantee"] = "2", ["One"] = "1", ["Two"] = "2"
    };

    /// <summary>By Name form uses different Sort values than By Date.</summary>
    public static readonly Dictionary<string, string> SortOrderToRissByName = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "CAST([FileDate] as DATE) ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "CAST([FileDate] as DATE) DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    /// <summary>RISS Document Type values (option value -> display text). Use documentTypes with exact text e.g. ["DEED","MORTGAGE","RELEASE OF MORTGAGE"].</summary>
    public static readonly Dictionary<string, string> DocumentTypeOptions = new(StringComparer.OrdinalIgnoreCase)
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
    public static readonly Dictionary<string, string> SortOrderToRissByType = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "[FileDateSort] ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "[FileDateSort] DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    public static readonly Dictionary<string, string> IndexTypeToRissBookPage = new(StringComparer.OrdinalIgnoreCase)
    {
        ["All"] = "%", ["%"] = "%",
        ["DEED"] = "DBK", ["DBK"] = "DBK", ["Deeds"] = "DBK",
        ["MTG"] = "MBK", ["MBK"] = "MBK", ["Mortgages"] = "MBK", ["Mortgage"] = "MBK",
        ["PLT"] = "PBK", ["PBK"] = "PBK"
    };

    public static readonly Dictionary<string, string> IndexTypeToRissPre1980 = new(StringComparer.OrdinalIgnoreCase)
    {
        ["All"] = "%", ["%"] = "%",
        ["PDE"] = "PDE", ["DEED"] = "PDE", ["Deeds"] = "PDE",
        ["PMT"] = "PMT", ["MTG"] = "PMT", ["Mortgages"] = "PMT", ["Mortgage"] = "PMT"
    };

    public static readonly Dictionary<string, string> IndexTypeToRiss = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Deeds"] = "DEE", ["Mortgages"] = "MTG", ["Partnerships"] = "PTR",
        ["UCC"] = "FIN", ["Service Discharge"] = "DIS", ["Veteran Graves"] = "VET"
    };

    public static readonly Dictionary<string, string> SortOrderToRiss = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MATCH ASC"] = "[Match] ASC", ["FILE DATE ASC"] = "[FileDateSort] ASC",
        ["FILE NUMBER ASC"] = "[InstrumentNum] ASC", ["TYPE ASC"] = "[DocDesc] ASC",
        ["MATCH DESC"] = "[Match] DESC", ["FILE DATE DESC"] = "[FileDateSort] DESC",
        ["FILE NUMBER DESC"] = "[InstrumentNum] DESC", ["TYPE DESC"] = "[DocDesc] DESC"
    };

    public static async Task FillCaptchaAsync(IPage page, string captchaCode)
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

    public static async Task ClickSidebarTabAsync(IPage page, string searchMode)
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
        var tab = page.GetByText(tabText, new() { Exact = false });
        await tab.First.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 15000 });
        await tab.First.ScrollIntoViewIfNeededAsync();
        await tab.First.ClickAsync();
        await page.WaitForTimeoutAsync(500);
    }

    /// <summary>By Name form: LastName*, FirstName, Side, IndexType*, StartDate, EndDate, Sort. No captcha.</summary>
    public static async Task FillByNameFormBatchAsync(IPage page, ActorInput input)
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

    public static string NormalizeNameSearch(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return "BeginWith";
        var t = v.Trim();
        if (t.Equals("Contains", StringComparison.OrdinalIgnoreCase)) return "Contains";
        if (t.Equals("Exactly", StringComparison.OrdinalIgnoreCase)) return "Exactly";
        if (t.Contains("Begin", StringComparison.OrdinalIgnoreCase)) return "BeginWith";
        return "BeginWith";
    }

    /// <summary>By Type form: Document Type*, Include Federal Lien, Index Type*, LastName, FirstName, Side, StartDate, EndDate, Sort.</summary>
    public static async Task FillByTypeForm(IPage page, ActorInput input)
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
    public static async Task FillByMunicipalityForm(IPage page, ActorInput input)
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
    public static async Task FillBySubdivisionForm(IPage page, ActorInput input)
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
    public static async Task FillBySTRForm(IPage page, ActorInput input)
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
    public static async Task FillByInstrumentForm(IPage page, ActorInput input)
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

    /// <summary>By Book/Page: Book*, Page*, Thru (optional), IndexType (All/DEED/MTG/PLT).</summary>
    public static async Task FillByBookPageForm(IPage page, ActorInput input)
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
    public static async Task FillByFicheForm(IPage page, ActorInput input)
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

    /// <summary>By Pre-1980: PreYear (1971-1979), PreNumber* (required), IndexType (All/PDE/PMT).</summary>
    public static async Task FillByPre1980Form(IPage page, ActorInput input)
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

    public static async Task FillByDateFormBatchAsync(IPage page, ActorInput input, bool fillCaptcha)
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

    public static async Task FillCommonFields(IPage page, ActorInput input)
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

    public static async Task TrySelectChosenAsync(IPage page, string labelOrId, string? value)
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

    public static async Task TryFillAsync(IPage page, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var el = page.GetByLabel(label).Or(page.Locator($"input[id*='{label.Replace(" ", "")}'], input[name*='{label.Replace(" ", "")}']").First);
        if (await el.CountAsync() > 0)
            await el.First.FillAsync(value);
    }

    public static async Task TrySelectAsync(IPage page, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        var sel = page.GetByLabel(label).Or(page.Locator("select[id*='Side'], select[name*='Side']").First);
        if (await sel.CountAsync() > 0)
            await sel.First.SelectOptionAsync(new[] { value });
    }

    public static async Task TryFillOrSelectAsync(IPage page, string label, string? value)
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
}
