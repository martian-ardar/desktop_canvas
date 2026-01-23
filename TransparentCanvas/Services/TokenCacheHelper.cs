using System.IO;
using Microsoft.Identity.Client;

namespace TransparentCanvas.Services;

/// <summary>
/// Token 缓存助手类 - 用于持久化存储认证 token
/// </summary>
public static class TokenCacheHelper
{
    private static readonly string CacheFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TransparentCanvas",
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
                    // 如果保存失败，忽略
                }
            }
        }
    }

    public static void ClearCache()
    {
        if (File.Exists(CacheFilePath))
        {
            File.Delete(CacheFilePath);
        }
    }
}
