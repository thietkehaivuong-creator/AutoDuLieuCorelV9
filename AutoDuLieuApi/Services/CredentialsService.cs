namespace AutoDuLieuApi.Services;

using AutoDuLieuApi.Models;
using System.Text.Json;

public interface ICredentialsService
{
    Task<CredentialsSettings> GetCredentialsAsync();
    Task SaveCredentialsAsync(string username, string password);
    Task<bool> ValidateCredentialsAsync(string username, string password);
}

public class CredentialsService : ICredentialsService
{
    private readonly string _credentialsFile;
    private readonly ILogger<CredentialsService> _logger;

    public CredentialsService(ILogger<CredentialsService> logger)
    {
        _logger = logger;
        var dataPath = Path.Combine(Directory.GetCurrentDirectory(), "data", "config");
        Directory.CreateDirectory(dataPath);
        _credentialsFile = Path.Combine(dataPath, "credentials.json");
    }

    public async Task<CredentialsSettings> GetCredentialsAsync()
    {
        try
        {
            if (!File.Exists(_credentialsFile))
                return new CredentialsSettings();

            var json = await File.ReadAllTextAsync(_credentialsFile);
            var credentials = JsonSerializer.Deserialize<CredentialsSettings>(json);
            return credentials ?? new CredentialsSettings();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading credentials: {ex.Message}");
            return new CredentialsSettings();
        }
    }

    public async Task SaveCredentialsAsync(string username, string password)
    {
        try
        {
            var credentials = new CredentialsSettings
            {
                Username = username,
                Password = password
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(credentials, options);

            Directory.CreateDirectory(Path.GetDirectoryName(_credentialsFile)!);
            await File.WriteAllTextAsync(_credentialsFile, json);

            _logger.LogInformation("Credentials saved successfully");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error saving credentials: {ex.Message}");
            throw;
        }
    }

    public async Task<bool> ValidateCredentialsAsync(string username, string password)
    {
        var credentials = await GetCredentialsAsync();
        return !string.IsNullOrWhiteSpace(credentials.Username) && 
               !string.IsNullOrWhiteSpace(credentials.Password);
    }
}