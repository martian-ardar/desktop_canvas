using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

// 配置类
class AppConfig
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

class AzureAdConfig
{
    public string ClientId { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
}

class OneNoteConfig
{
    public string NotebookName { get; set; } = string.Empty;
    public string SectionName { get; set; } = string.Empty;
}

// Token 缓存助手类
static class TokenCacheHelper
{
    private static readonly string CacheFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OneNoteApp",
        "msal_cache.dat"
    );

    private static readonly object FileLock = new object();

    public static void EnableSerialization(ITokenCache tokenCache)
    {
        tokenCache.SetBeforeAccess(BeforeAccessNotification);
        tokenCache.SetAfterAccess(AfterAccessNotification);
    }

    private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
    {
        lock (FileLock)
        {
            // 从文件加载缓存
            if (File.Exists(CacheFilePath))
            {
                try
                {
                    byte[] data = File.ReadAllBytes(CacheFilePath);
                    args.TokenCache.DeserializeMsalV3(data);
                }
                catch
                {
                    // 如果缓存损坏，忽略并重新登录
                }
            }
        }
    }

    private static void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
        // 如果缓存状态已改变，保存到文件
        if (args.HasStateChanged)
        {
            lock (FileLock)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(CacheFilePath)!);
                    byte[] data = args.TokenCache.SerializeMsalV3();
                    File.WriteAllBytes(CacheFilePath, data);
                }
                catch
                {
                    // 如果保存失败，忽略（下次会重新登录）
                }
            }
        }
    }

    public static void ClearCache()
    {
        if (File.Exists(CacheFilePath))
        {
            File.Delete(CacheFilePath);
            Console.WriteLine("✓ 缓存已清除");
        }
    }
}

class Program
{
    static async Task Main(string[] args)
    {
        // 检查是否需要清除缓存
        if (args.Length > 0 && args[0] == "--clear-cache")
        {
            TokenCacheHelper.ClearCache();
            return;
        }

        // 加载配置
        AppConfig config;
        try
        {
            config = AppConfig.Load();
            Console.WriteLine("✓ 配置文件加载成功");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"加载配置文件失败: {ex.Message}");
            return;
        }

        var scopes = new[] { "Notes.ReadWrite" };

        var app = PublicClientApplicationBuilder
            .Create(config.AzureAd.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{config.AzureAd.TenantId}")
            .WithDefaultRedirectUri()
            .Build();

        // 启用 token 缓存
        TokenCacheHelper.EnableSerialization(app.UserTokenCache);

        AuthenticationResult result;

        try
        {
            // 1. 尝试静默获取 token（从缓存）
            var accounts = await app.GetAccountsAsync();
            if (accounts.Any())
            {
                Console.WriteLine("正在使用缓存的凭据登录...");
                result = await app.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
                Console.WriteLine("✓ 登录成功（使用缓存）\n");
            }
            else
            {
                throw new MsalUiRequiredException("no_cached_account", "需要重新登录");
            }
        }
        catch (MsalUiRequiredException)
        {
            // 2. 如果静默获取失败，使用 Device Code Flow
            Console.WriteLine("需要重新登录...");
            result = await app.AcquireTokenWithDeviceCode(
                scopes,
                code =>
                {
                    Console.WriteLine(code.Message);
                    return Task.CompletedTask;
                }).ExecuteAsync();
            Console.WriteLine("✓ 登录成功\n");
        }

        var accessToken = result.AccessToken;

        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

        // 1. 获取笔记本列表，查找配置的笔记本
        var notebooksResponse = await http.GetAsync(
            "https://graph.microsoft.com/v1.0/me/onenote/notebooks");
        var notebooksJson = await notebooksResponse.Content.ReadAsStringAsync();
        
        Console.WriteLine($"正在查找笔记本 '{config.OneNote.NotebookName}'...");
        var notebooksDoc = JsonDocument.Parse(notebooksJson);
        var notebook = notebooksDoc.RootElement
            .GetProperty("value")
            .EnumerateArray()
            .FirstOrDefault(n => n.GetProperty("displayName").GetString() == config.OneNote.NotebookName);
        
        if (notebook.ValueKind == JsonValueKind.Undefined)
        {
            Console.WriteLine($"未找到笔记本 '{config.OneNote.NotebookName}'");
            return;
        }
        
        var notebookId = notebook.GetProperty("id").GetString();
        Console.WriteLine($"找到笔记本: {notebook.GetProperty("displayName").GetString()} (ID: {notebookId})");

        // 2. 检查是否有分区组
        Console.WriteLine("\n正在检查分区组...");
        var sectionGroupsResponse = await http.GetAsync(
            $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sectionGroups");
        var sectionGroupsJson = await sectionGroupsResponse.Content.ReadAsStringAsync();
        var sectionGroupsDoc = JsonDocument.Parse(sectionGroupsJson);
        var sectionGroups = sectionGroupsDoc.RootElement.GetProperty("value").EnumerateArray().ToList();
        
        if (sectionGroups.Count > 0)
        {
            Console.WriteLine($"找到 {sectionGroups.Count} 个分区组:");
            foreach (var sg in sectionGroups)
            {
                Console.WriteLine($"  - {sg.GetProperty("displayName").GetString()}");
            }
        }
        else
        {
            Console.WriteLine("没有分区组");
        }

        // 3. 获取该笔记本下的分区列表，查找 "from_wallpaper"
        var sectionsResponse = await http.GetAsync(
            $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sections");
        var sectionsJson = await sectionsResponse.Content.ReadAsStringAsync();
        
        Console.WriteLine("\n正在查找分区...");
        var sectionsDoc = JsonDocument.Parse(sectionsJson);
        var sections = sectionsDoc.RootElement.GetProperty("value").EnumerateArray().ToList();
        
        // 显示所有找到的分区
        Console.WriteLine($"当前笔记本中的分区列表 (共 {sections.Count} 个):");
        foreach (var s in sections)
        {
            var displayName = s.GetProperty("displayName").GetString();
            var id = s.GetProperty("id").GetString();
            Console.WriteLine($"  - 名称: {displayName}");
            Console.WriteLine($"    ID: {id}");
            
            // 如果有 self URL，也显示出来
            if (s.TryGetProperty("self", out var selfProp))
            {
                Console.WriteLine($"    URL: {selfProp.GetString()}");
            }
        }
        
        var section = sections.FirstOrDefault(s => s.GetProperty("displayName").GetString() == config.OneNote.SectionName);
        
        string sectionId;
        if (section.ValueKind == JsonValueKind.Undefined)
        {
            Console.WriteLine($"\n未找到分区 '{config.OneNote.SectionName}'，正在创建新分区...");
            
            // 创建新分区
            var newSectionJson = JsonSerializer.Serialize(new { displayName = config.OneNote.SectionName });
            var newSectionContent = new StringContent(newSectionJson, Encoding.UTF8, "application/json");
            
            var createSectionResponse = await http.PostAsync(
                $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sections",
                newSectionContent
            );
            
            if (!createSectionResponse.IsSuccessStatusCode)
            {
                Console.WriteLine($"创建分区失败: {createSectionResponse.StatusCode}");
                Console.WriteLine(await createSectionResponse.Content.ReadAsStringAsync());
                return;
            }
            
            var newSectionJson2 = await createSectionResponse.Content.ReadAsStringAsync();
            var newSectionDoc = JsonDocument.Parse(newSectionJson2);
            sectionId = newSectionDoc.RootElement.GetProperty("id").GetString()!;
            Console.WriteLine($"✓ 成功创建分区: {config.OneNote.SectionName} (ID: {sectionId})");
        }
        else
        {
            sectionId = section.GetProperty("id").GetString()!;
            Console.WriteLine($"找到分区: {section.GetProperty("displayName").GetString()} (ID: {sectionId})");
        }

        // 4. 在指定分区下创建页面，标题为 "wall_1"
        var html = $@"<!DOCTYPE html>
<html>
  <head>
    <title>wall_3</title>
  </head>
  <body>
    <h1>wall_3</h1>
    <p>这是通过 Microsoft Graph API 创建的页面。</p>
    <p>创建时间: {DateTime.Now}</p>
  </body>
</html>";

        var content = new StringContent(html, Encoding.UTF8, "application/xhtml+xml");

        Console.WriteLine("正在创建页面...");
        var response = await http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/onenote/sections/{sectionId}/pages",
            content
        );

        var responseText = await response.Content.ReadAsStringAsync();

        Console.WriteLine($"Status: {response.StatusCode}");
        Console.WriteLine(responseText);
        
        if (response.IsSuccessStatusCode)
        {
            Console.WriteLine($"\n✓ 页面已成功创建在 '{config.OneNote.NotebookName} / {config.OneNote.SectionName}' 下");
            Console.WriteLine("\n提示：下次运行将自动使用缓存的凭据，无需重新登录");
            Console.WriteLine("如需重新登录，请运行: dotnet run -- --clear-cache");
        }
    }
}