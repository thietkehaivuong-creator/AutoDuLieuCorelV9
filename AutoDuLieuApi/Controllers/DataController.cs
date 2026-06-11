namespace AutoDuLieuApi.Controllers;

using AutoDuLieuApi.Models;
using AutoDuLieuApi.Services;
using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/[controller]")]
public class DataController : ControllerBase
{
    private readonly IDataProcessingService _dataService;
    private readonly ICredentialsService _credentialsService;
    private readonly ILogger<DataController> _logger;

    public DataController(
        IDataProcessingService dataService,
        ICredentialsService credentialsService,
        ILogger<DataController> logger)
    {
        _dataService = dataService;
        _credentialsService = credentialsService;
        _logger = logger;
    }

    [HttpGet("search/{pfi}")]
    public async Task<ActionResult<SearchResponse>> SearchByPfi(string pfi)
    {
        _logger.LogInformation($"Search PFI: {pfi}");
        
        var result = await _dataService.SearchByPfiAsync(pfi);
        
        if (result == null)
            return Ok(new SearchResponse 
            { 
                Success = false, 
                Message = $"PFI '{pfi}' không tìm thấy" 
            });

        return Ok(new SearchResponse
        {
            Success = true,
            Message = "Tìm thấy dữ liệu",
            Data = result
        });
    }

    [HttpGet("all")]
    public async Task<ActionResult<List<DataRow>>> GetAllData()
    {
        var data = await _dataService.GetAllDataAsync();
        return Ok(data);
    }

    [HttpPost("format-date")]
    public ActionResult<DateFormatResponse> FormatDate([FromBody] DateFormatRequest request)
    {
        try
        {
            var formatted = _dataService.FormatDate(request.InputDate, request.TargetFormat);
            return Ok(new DateFormatResponse
            {
                Success = true,
                FormattedDate = formatted
            });
        }
        catch (Exception ex)
        {
            return Ok(new DateFormatResponse
            {
                Success = false,
                Error = ex.Message
            });
        }
    }

    [HttpGet("fao/{code}")]
    public ActionResult<object> GetFAOName(string code)
    {
        var name = _dataService.LookupFAOName(code);
        return Ok(new { code, name });
    }

    [HttpGet("ppdb/{code}")]
    public ActionResult<object> GetPPDBName(string code)
    {
        var name = _dataService.LookupPPDBName(code);
        return Ok(new { code, name });
    }
}

[ApiController]
[Route("api/[controller]")]
public class SettingsController : ControllerBase
{
    private readonly ICredentialsService _credentialsService;
    private readonly ILogger<SettingsController> _logger;

    public SettingsController(
        ICredentialsService credentialsService,
        ILogger<SettingsController> logger)
    {
        _credentialsService = credentialsService;
        _logger = logger;
    }

    [HttpGet("credentials")]
    public async Task<ActionResult<SettingsResponse>> GetCredentials()
    {
        var credentials = await _credentialsService.GetCredentialsAsync();
        return Ok(new SettingsResponse
        {
            Success = true,
            Message = "Đã lấy thông tin",
            Data = new CredentialsSettings 
            { 
                Username = credentials.Username,
                Password = string.IsNullOrEmpty(credentials.Password) ? "" : "***"
            }
        });
    }

    [HttpPost("credentials")]
    public async Task<ActionResult<SettingsResponse>> SaveCredentials([FromBody] CredentialsSettings request)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(request.Username) || 
                string.IsNullOrWhiteSpace(request.Password))
            {
                return Ok(new SettingsResponse
                {
                    Success = false,
                    Message = "Username và Password không được để trống"
                });
            }

            await _credentialsService.SaveCredentialsAsync(request.Username, request.Password);

            return Ok(new SettingsResponse
            {
                Success = true,
                Message = "Đã lưu thông tin đăng nhập",
                Data = new CredentialsSettings { Username = request.Username }
            });
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error saving credentials: {ex.Message}");
            return Ok(new SettingsResponse
            {
                Success = false,
                Message = $"Lỗi: {ex.Message}"
            });
        }
    }

    [HttpGet("status")]
    public async Task<ActionResult<object>> GetStatus()
    {
        var credentials = await _credentialsService.GetCredentialsAsync();
        var hasCredentials = !string.IsNullOrWhiteSpace(credentials.Username);

        return Ok(new
        {
            success = true,
            hasCredentials = hasCredentials,
            username = hasCredentials ? credentials.Username : null,
            message = hasCredentials ? "Đã có thông tin đăng nhập" : "Chưa cấu hình thông tin đăng nhập"
        });
    }
}