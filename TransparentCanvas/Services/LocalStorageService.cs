using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using TransparentCanvas.Models;

namespace TransparentCanvas.Services;

/// <summary>
/// 本地存储服务 - 将日程提醒类笔记保存到本地磁盘
/// </summary>
public class LocalStorageService
{
    private static readonly string StorageRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TransparentCanvas", "pages");

    public LocalStorageService()
    {
        Directory.CreateDirectory(StorageRoot);
    }

    /// <summary>
    /// 保存页面到本地（先保存内存中的状态到 canvas，再调用此方法）
    /// </summary>
    public void SavePage(PageData page, InkCanvas canvas)
    {
        try
        {
            // 先将 canvas 当前内容保存到 page 内存
            page.SaveFromCanvas(canvas);
            SavePageData(page);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"保存页面失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 保存 PageData（从内存中的数据保存，不需要 canvas）
    /// </summary>
    public void SavePageData(PageData page)
    {
        try
        {
            var pageDir = Path.Combine(StorageRoot, page.StorageId);
            Directory.CreateDirectory(pageDir);

            // 1. 保存笔画 (ISF 格式)
            var strokesPath = Path.Combine(pageDir, "strokes.isf");
            if (page.Strokes.Count > 0)
            {
                using var fs = new FileStream(strokesPath, FileMode.Create);
                page.Strokes.Save(fs);
            }
            else if (File.Exists(strokesPath))
            {
                File.Delete(strokesPath);
            }

            // 2. 收集子元素信息
            var childrenData = new List<ChildElementJson>();
            int imageIndex = 0;
            foreach (var child in page.Children)
            {
                if (child.Element is TextBlock textBlock)
                {
                    var color = "#000000";
                    if (textBlock.Foreground is SolidColorBrush brush)
                    {
                        color = $"#{brush.Color.A:X2}{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                    }

                    childrenData.Add(new ChildElementJson
                    {
                        Type = "text",
                        Text = textBlock.Text,
                        Left = child.Left,
                        Top = child.Top,
                        FontSize = textBlock.FontSize,
                        ForegroundColor = color
                    });
                }
                else if (child.Element is System.Windows.Controls.Image image && image.Source is BitmapSource bitmapSource)
                {
                    // 保存图片为 PNG 文件
                    var imageName = $"image_{imageIndex++}.png";
                    var imagePath = Path.Combine(pageDir, imageName);
                    SaveBitmapToPng(bitmapSource, imagePath);

                    childrenData.Add(new ChildElementJson
                    {
                        Type = "image",
                        ImageFile = imageName,
                        Left = child.Left,
                        Top = child.Top,
                        MaxWidth = image.MaxWidth,
                        MaxHeight = image.MaxHeight
                    });
                }
            }

            // 3. 保存元数据 JSON
            var meta = new PageMetaJson
            {
                StorageId = page.StorageId,
                Title = page.Title,
                NoteType = page.NoteType.ToString(),
                TargetTime = page.TargetTime,
                HasReminded = page.HasReminded,
                CreatedAt = page.CreatedAt,
                ModifiedAt = DateTime.Now,
                Children = childrenData
            };

            var metaPath = Path.Combine(pageDir, "meta.json");
            var json = JsonSerializer.Serialize(meta, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(metaPath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"保存页面数据失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 加载所有已保存的页面
    /// </summary>
    public List<PageData> LoadAllPages()
    {
        var pages = new List<PageData>();

        if (!Directory.Exists(StorageRoot))
            return pages;

        foreach (var pageDir in Directory.GetDirectories(StorageRoot))
        {
            try
            {
                var page = LoadPage(pageDir);
                if (page != null)
                {
                    pages.Add(page);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载页面失败 ({pageDir}): {ex.Message}");
            }
        }

        return pages;
    }

    /// <summary>
    /// 加载单个页面
    /// </summary>
    private PageData? LoadPage(string pageDir)
    {
        var metaPath = Path.Combine(pageDir, "meta.json");
        if (!File.Exists(metaPath))
            return null;

        var json = File.ReadAllText(metaPath);
        var meta = JsonSerializer.Deserialize<PageMetaJson>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (meta == null)
            return null;

        var noteType = Enum.TryParse<NoteType>(meta.NoteType, out var nt) ? nt : NoteType.Normal;

        var page = new PageData(
            storageId: meta.StorageId,
            title: meta.Title,
            noteType: noteType,
            createdAt: meta.CreatedAt,
            modifiedAt: meta.ModifiedAt,
            targetTime: meta.TargetTime,
            hasReminded: meta.HasReminded
        );

        // 加载笔画
        var strokesPath = Path.Combine(pageDir, "strokes.isf");
        if (File.Exists(strokesPath))
        {
            using var fs = new FileStream(strokesPath, FileMode.Open);
            page.Strokes = new StrokeCollection(fs);
        }

        // 加载子元素
        if (meta.Children != null)
        {
            foreach (var childJson in meta.Children)
            {
                UIElement? element = null;
                
                if (childJson.Type == "text" && childJson.Text != null)
                {
                    var textBlock = new TextBlock
                    {
                        Text = childJson.Text,
                        FontSize = childJson.FontSize > 0 ? childJson.FontSize : 14,
                        Background = Brushes.Transparent,
                        Padding = new Thickness(4),
                        TextWrapping = TextWrapping.Wrap,
                        MaxWidth = 400,
                        Cursor = System.Windows.Input.Cursors.Hand
                    };

                    if (!string.IsNullOrEmpty(childJson.ForegroundColor))
                    {
                        try
                        {
                            textBlock.Foreground = new SolidColorBrush(
                                (Color)ColorConverter.ConvertFromString(childJson.ForegroundColor));
                        }
                        catch
                        {
                            textBlock.Foreground = Brushes.Black;
                        }
                    }

                    element = textBlock;
                }
                else if (childJson.Type == "image" && childJson.ImageFile != null)
                {
                    var imagePath = Path.Combine(pageDir, childJson.ImageFile);
                    if (File.Exists(imagePath))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                        bitmap.EndInit();
                        bitmap.Freeze();

                        element = new System.Windows.Controls.Image
                        {
                            Source = bitmap,
                            Stretch = Stretch.Uniform,
                            MaxWidth = childJson.MaxWidth > 0 ? childJson.MaxWidth : 400,
                            MaxHeight = childJson.MaxHeight > 0 ? childJson.MaxHeight : 300
                        };
                    }
                }

                if (element != null)
                {
                    page.Children.Add(new ChildElementData
                    {
                        Element = element,
                        Left = childJson.Left,
                        Top = childJson.Top
                    });
                }
            }
        }

        return page;
    }

    /// <summary>
    /// 删除页面的本地存储
    /// </summary>
    public void DeletePage(PageData page)
    {
        try
        {
            var pageDir = Path.Combine(StorageRoot, page.StorageId);
            if (Directory.Exists(pageDir))
            {
                Directory.Delete(pageDir, true);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"删除页面存储失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 保存 BitmapSource 为 PNG 文件
    /// </summary>
    private void SaveBitmapToPng(BitmapSource bitmap, string path)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using var fs = new FileStream(path, FileMode.Create);
        encoder.Save(fs);
    }
}

#region JSON 序列化模型

internal class PageMetaJson
{
    public string StorageId { get; set; } = "";
    public string Title { get; set; } = "";
    public string NoteType { get; set; } = "Normal";
    public DateTime? TargetTime { get; set; }
    public bool HasReminded { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime ModifiedAt { get; set; }
    public List<ChildElementJson>? Children { get; set; }
}

internal class ChildElementJson
{
    public string Type { get; set; } = ""; // "text" or "image"
    public string? Text { get; set; }
    public string? ImageFile { get; set; }
    public double Left { get; set; }
    public double Top { get; set; }
    public double FontSize { get; set; }
    public string? ForegroundColor { get; set; }
    public double MaxWidth { get; set; }
    public double MaxHeight { get; set; }
}

#endregion