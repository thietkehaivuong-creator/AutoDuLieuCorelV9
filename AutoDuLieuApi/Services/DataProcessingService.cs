namespace AutoDuLieuApi.Services;

using AutoDuLieuApi.Models;

public interface IDataProcessingService
{
    Task<DataRow?> SearchByPfiAsync(string pfi);
    Task<List<DataRow>> GetAllDataAsync();
    string FormatDate(string input, string format);
    List<string> SplitFAO(string faoText);
    string LookupFAOName(string code);
    string LookupPPDBName(string code);
}

public class DataProcessingService : IDataProcessingService
{
    private readonly IConfiguration _config;
    private readonly ILogger<DataProcessingService> _logger;
    private List<DataRow> _cachedData = new();
    private Dictionary<string, string> _faoMapping = new();
    private Dictionary<string, string> _ppdbMapping = new();
    private DateTime _lastDataLoad = DateTime.MinValue;

    public DataProcessingService(IConfiguration config, ILogger<DataProcessingService> logger)
    {
        _config = config;
        _logger = logger;
        InitializeMappings();
    }

    private void InitializeMappings()
    {
        _faoMapping = new Dictionary<string, string>
        {
            { "18", "Arctic Ocean" },
            { "21", "Northwestern Atlantic Ocean" },
            { "27", "Northeastern Atlantic Ocean" },
            { "31", "Western Central Atlantic Ocean" },
            { "34", "Eastern Central Atlantic Ocean" },
            { "37", "Mediterranean and Black Sea" },
            { "41", "Southwestern Atlantic Ocean" },
            { "47", "Southeastern Atlantic Ocean" },
            { "48", "Antarctic Atlantic Ocean" },
            { "51", "Western Indian Ocean" },
            { "57", "Eastern Indian Ocean" },
            { "58", "Antarctic and Southern Indian Ocean" },
            { "61", "Northwestern Pacific Ocean" },
            { "67", "Northeastern Pacific Ocean" },
            { "71", "Western Central Pacific Ocean" },
            { "77", "Eastern Central Pacific Ocean" },
            { "81", "Southwestern Pacific Ocean" },
            { "87", "Southeastern Pacific Ocean" },
            { "88", "Antarctic Pacific Ocean" }
        };

        _ppdbMapping = new Dictionary<string, string>
        {
            { "PS", "Purse seine" },
            { "LHP", "Hook and line" },
            { "HNL", "Hook and line" },
            { "HL", "Hand line" },
            { "LL", "Long line" },
            { "PL", "Pole and line" },
            { "TR", "Trawl" },
            { "GN", "Gillnet" },
            { "DN", "Dip net" },
            { "SN", "Surrounding net" },
            { "RN", "Ring net" }
        };
    }

    public async Task<DataRow?> SearchByPfiAsync(string pfi)
    {
        if (string.IsNullOrWhiteSpace(pfi))
            return null;

        await EnsureDataLoadedAsync();

        var result = _cachedData.FirstOrDefault(x => 
            x.PFI.Equals(pfi.Trim(), StringComparison.OrdinalIgnoreCase));

        _logger.LogInformation($"Search PFI: {pfi} - Found: {(result != null ? "Yes" : "No")}");
        return result;
    }

    public async Task<List<DataRow>> GetAllDataAsync()
    {
        await EnsureDataLoadedAsync();
        return _cachedData;
    }

    private async Task EnsureDataLoadedAsync()
    {
        if (_lastDataLoad > DateTime.UtcNow.AddMinutes(-5))
            return;

        try
        {
            _cachedData = GetDemoData();
            _lastDataLoad = DateTime.UtcNow;
            _logger.LogInformation($"Data loaded: {_cachedData.Count} rows");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error loading data: {ex.Message}");
            _cachedData = GetDemoData();
        }
    }

    private List<DataRow> GetDemoData()
    {
        return new List<DataRow>
        {
            new DataRow
            {
                Customer = "ABC SEAFOOD",
                PFI = "24001",
                Lo = "A",
                Description = "TUNA LOIN",
                Sizes = "2KG",
                Qty = "100 CTN",
                SPEC = "SPEC A",
                Xn_Tem_Hang = "02/06/2026",
                Ng_DG_tu = "02/06/2028",
                Ng_DG = "01/06/2026",
                Best_Before = "01/06/2026",
                Nguon_NL = "03/06/2026",
                FAO = "71;77",
                haTGDB = "02/06/2028",
                PP_DB = "PS",
                Lot_No = "LOT-ABC-001",
                Remark = "2 FAO demo",
                NLT_ve = "05/06/2026",
                TGDB_NK = "10 x 2KG",
                Packing = "10 x 2KG"
            },
            new DataRow
            {
                Customer = "ABC SEAFOOD",
                PFI = "24002",
                Lo = "B",
                Description = "TUNA STEAK",
                Sizes = "1KG",
                Qty = "80 CTN",
                SPEC = "SPEC B",
                Xn_Tem_Hang = "10/06/2026",
                Ng_DG_tu = "10/06/2028",
                Ng_DG = "09/06/2026",
                Best_Before = "09/06/2026",
                Nguon_NL = "11/06/2026",
                FAO = "71",
                haTGDB = "10/06/2028",
                PP_DB = "LHP",
                Lot_No = "LOT-ABC-002",
                Remark = "1 FAO demo",
                NLT_ve = "12/06/2026",
                TGDB_NK = "20 x 1KG",
                Packing = "20 x 1KG"
            },
            new DataRow
            {
                Customer = "XYZ IMPORT",
                PFI = "24003",
                Lo = "C",
                Description = "SALMON PORTION",
                Sizes = "500G",
                Qty = "120 CTN",
                SPEC = "SPEC X",
                Xn_Tem_Hang = "15/06/2026",
                Ng_DG_tu = "15/06/2028",
                Ng_DG = "14/06/2026",
                Best_Before = "14/06/2026",
                Nguon_NL = "16/06/2026",
                FAO = "27;51",
                haTGDB = "15/06/2028",
                PP_DB = "HL",
                Lot_No = "LOT-XYZ-001",
                Remark = "PFI trung voi ABC",
                NLT_ve = "17/06/2026",
                TGDB_NK = "20 x 500G",
                Packing = "20 x 500G"
            }
        };
    }

    public string FormatDate(string input, string format)
    {
        if (string.IsNullOrWhiteSpace(input))
            return "";

        try
        {
            var date = ParseDateFlexible(input);
            return format.ToLower() switch
            {
                "dd/mm/yyyy" => date.ToString("dd/MM/yyyy"),
                "d/m/yyyy" => date.ToString("d/M/yyyy"),
                "mm/dd/yyyy" => date.ToString("MM/dd/yyyy"),
                "m/d/yyyy" => date.ToString("M/d/yyyy"),
                "yyyy/mm/dd" => date.ToString("yyyy/MM/dd"),
                "yyyy/m/d" => date.ToString("yyyy/M/d"),
                "dd-mm-yyyy" => date.ToString("dd-MM-yyyy"),
                "mm-dd-yyyy" => date.ToString("MM-dd-yyyy"),
                "yyyy-mm-dd" => date.ToString("yyyy-MM-dd"),
                "dd.mm.yyyy" => date.ToString("dd.MM.yyyy"),
                "mm.yyyy" => date.ToString("MM.yyyy"),
                "yyyy.mm.dd" => date.ToString("yyyy.MM.dd"),
                "mm/yyyy" => date.ToString("MM/yyyy"),
                "m/yyyy" => date.ToString("M/yyyy"),
                "yyyy/mm" => date.ToString("yyyy/MM"),
                "yyyy-mm" => date.ToString("yyyy-MM"),
                "mmm yyyy" => date.ToString("MMM yyyy"),
                "mmmm yyyy" => date.ToString("MMMM yyyy"),
                "dd mmm yyyy" => date.ToString("dd MMM yyyy"),
                "mmm dd, yyyy" => date.ToString("MMM dd, yyyy"),
                _ => date.ToString(format)
            };
        }
        catch (Exception ex)
        {
            _logger.LogError($"Date formatting error: {ex.Message}");
            return input;
        }
    }

    public List<string> SplitFAO(string faoText)
    {
        if (string.IsNullOrWhiteSpace(faoText))
            return new List<string>();

        return faoText
            .Replace("FAO", "")
            .Replace("+", ";")
            .Replace("/", ";")
            .Replace(",", ";")
            .Replace(" AND ", ";")
            .Split(';')
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrEmpty(x))
            .ToList();
    }

    public string LookupFAOName(string code)
    {
        return _faoMapping.TryGetValue(code?.Trim() ?? "", out var name) ? name : code ?? "";
    }

    public string LookupPPDBName(string code)
    {
        return _ppdbMapping.TryGetValue(code?.Trim() ?? "", out var name) ? name : code ?? "";
    }

    private DateTime ParseDateFlexible(string input)
    {
        input = input?.Trim() ?? "";
        input = input.Replace("-", "/").Replace(".", "/");

        if (string.IsNullOrWhiteSpace(input))
            return DateTime.Now;

        var parts = input.Split('/');
        if (parts.Length < 2 || parts.Length > 3)
            return DateTime.Parse(input);

        if (parts.Length == 2)
        {
            var month = int.Parse(parts[0]);
            var year = int.Parse(parts[1]);
            if (year < 100) year += 2000;
            return new DateTime(year, month, 1);
        }

        int day, month, year;
        if (parts[0].Length == 4)
        {
            year = int.Parse(parts[0]);
            month = int.Parse(parts[1]);
            day = int.Parse(parts[2]);
        }
        else
        {
            day = int.Parse(parts[0]);
            month = int.Parse(parts[1]);
            year = int.Parse(parts[2]);
        }

        if (year < 100) year += 2000;
        return new DateTime(year, month, day);
    }
}