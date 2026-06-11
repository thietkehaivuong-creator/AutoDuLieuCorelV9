namespace AutoDuLieuApi.Models;

public class DataRow
{
    public string Customer { get; set; } = "";
    public string PFI { get; set; } = "";
    public string Lo { get; set; } = "";
    public string Description { get; set; } = "";
    public string Sizes { get; set; } = "";
    public string Qty { get; set; } = "";
    public string SPEC { get; set; } = "";
    public string Xn_Tem_Hang { get; set; } = "";
    public string Ng_DG_tu { get; set; } = "";
    public string Ng_DG { get; set; } = "";
    public string Best_Before { get; set; } = "";
    public string Nguon_NL { get; set; } = "";
    public string FAO { get; set; } = "";
    public string haTGDB { get; set; } = "";
    public string PP_DB { get; set; } = "";
    public string Lot_No { get; set; } = "";
    public string Remark { get; set; } = "";
    public string NLT_ve { get; set; } = "";
    public string TGDB_NK { get; set; } = "";
    public string Packing { get; set; } = "";

    public string GetValue(string fieldName)
    {
        return fieldName.ToLower() switch
        {
            "customer" => Customer,
            "pfi" => PFI,
            "lo" => Lo,
            "description" => Description,
            "sizes" => Sizes,
            "qty" => Qty,
            "spec" => SPEC,
            "xn_tem_hang" => Xn_Tem_Hang,
            "ng_dg_tu" => Ng_DG_tu,
            "ng_dg" => Ng_DG,
            "best_before" => Best_Before,
            "nguon_nl" => Nguon_NL,
            "fao" => FAO,
            "hatgdb" => haTGDB,
            "pp_db" => PP_DB,
            "lot_no" => Lot_No,
            "remark" => Remark,
            "nlt_ve" => NLT_ve,
            "tgdb_nk" => TGDB_NK,
            "packing" => Packing,
            _ => ""
        };
    }
}

public class SearchRequest
{
    public string Pfi { get; set; } = "";
}

public class SearchResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public DataRow Data { get; set; } = new();
}

public class DateFormatRequest
{
    public string InputDate { get; set; } = "";
    public string TargetFormat { get; set; } = "dd/mm/yyyy";
}

public class DateFormatResponse
{
    public string FormattedDate { get; set; } = "";
    public bool Success { get; set; }
    public string Error { get; set; } = "";
}

public class CredentialsSettings
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
}

public class SettingsResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public CredentialsSettings Data { get; set; } = new();
}