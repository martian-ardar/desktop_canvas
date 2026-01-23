using System;
using System.IO;
using System.Text.Json;

namespace TransparentCanvas.Config;

public class AppConfig
{
    public AzureAdConfig AzureAd { get; set; } = new();
    public OneNoteConfig OneNote { get; set; } = new();

    public static AppConfig Load()
    {
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        if (!File.Exists(configPath))
        {
            throw new FileNotFoundException($"配置文件未找到: {configPath}");
        }

        var json = File.ReadAllText(configPath);
        return JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? new AppConfig();
    }
}

public class AzureAdConfig
{
    public string ClientId { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
}

public class OneNoteConfig
{
    public string NotebookName { get; set; } = string.Empty;
    public string SectionName { get; set; } = string.Empty;
}

