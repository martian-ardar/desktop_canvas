using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Identity.Client;
using TransparentCanvas.Config;

namespace TransparentCanvas.Services;

/// <summary>
/// OneNote 服务 - 处理与 OneNote 的所有交互
/// </summary>
public class OneNoteService
{
    private readonly IPublicClientApplication _app;
    private readonly AppConfig _config;
    private readonly string[] _scopes = { "Notes.ReadWrite" };
    private string? _accessToken;

    public OneNoteService()
    {
        try
        {
            _config = AppConfig.Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"配置文件加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            throw;
        }

        _app = PublicClientApplicationBuilder
            .Create(_config.AzureAd.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{_config.AzureAd.TenantId}")
            .WithDefaultRedirectUri()
            .Build();

        TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
    }

    /// <summary>
    /// 获取访问令牌
    /// </summary>
    public async Task<bool> AuthenticateAsync()
    {
        try
        {
            AuthenticationResult result;
            var accounts = await _app.GetAccountsAsync();

            if (accounts.Any())
            {
                // 尝试静默获取 token
                result = await _app.AcquireTokenSilent(_scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {
                // 使用 Device Code Flow 进行交互式登录
                result = await _app.AcquireTokenWithDeviceCode(
                    _scopes,
                    code =>
                    {
                        // 自动打开浏览器并显示可复制的验证码
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            // 打开默认浏览器
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = code.VerificationUrl,
                                UseShellExecute = true
                            });
                            
                            // 显示可复制的验证码对话框
                            ShowDeviceCodeDialog(code.UserCode);
                        });
                        return Task.CompletedTask;
                    }).ExecuteAsync();
            }

            _accessToken = result.AccessToken;
            return true;
        }
        catch (MsalUiRequiredException)
        {
            // 需要用户交互
            try
            {
                var result = await _app.AcquireTokenWithDeviceCode(
                    _scopes,
                    code =>
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            // 打开默认浏览器
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = code.VerificationUrl,
                                UseShellExecute = true
                            });
                            
                            // 显示可复制的验证码对话框
                            ShowDeviceCodeDialog(code.UserCode);
                        });
                        return Task.CompletedTask;
                    }).ExecuteAsync();

                _accessToken = result.AccessToken;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"登录失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"认证失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    /// <summary>
    /// 显示设备代码对话框（可复制验证码）
    /// </summary>
    private void ShowDeviceCodeDialog(string userCode)
    {
        var window = new Window
        {
            Title = "OneNote 登录",
            Width = 400,
            Height = 200,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // 提示文字
        var label1 = new TextBlock
        {
            Text = "浏览器已打开，请在页面中输入以下验证码：",
            Margin = new Thickness(20, 20, 20, 10),
            FontSize = 14
        };
        Grid.SetRow(label1, 0);
        grid.Children.Add(label1);

        // 可选择复制的验证码文本框
        var codeTextBox = new TextBox
        {
            Text = userCode,
            FontSize = 24,
            FontWeight = FontWeights.Bold,
            TextAlignment = TextAlignment.Center,
            IsReadOnly = true,
            Margin = new Thickness(20, 10, 20, 10),
            Padding = new Thickness(10),
            Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0)),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC))
        };
        // 选中所有文本以便复制
        codeTextBox.Loaded += (s, e) => codeTextBox.SelectAll();
        Grid.SetRow(codeTextBox, 1);
        grid.Children.Add(codeTextBox);

        // 复制按钮
        var copyButton = new Button
        {
            Content = "复制验证码",
            Width = 100,
            Height = 30,
            Margin = new Thickness(20, 5, 20, 10)
        };
        copyButton.Click += (s, e) =>
        {
            Clipboard.SetText(userCode);
            copyButton.Content = "已复制 ✓";
        };
        Grid.SetRow(copyButton, 2);
        grid.Children.Add(copyButton);

        // 提示
        var label2 = new TextBlock
        {
            Text = "完成登录后此对话框将自动关闭",
            Margin = new Thickness(20, 5, 20, 15),
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)),
            HorizontalAlignment = HorizontalAlignment.Center
        };
        Grid.SetRow(label2, 3);
        grid.Children.Add(label2);

        window.Content = grid;
        window.Show();
    }

    /// <summary>
    /// 将画布内容保存到 OneNote
    /// </summary>
    public async Task<bool> SaveCanvasToOneNoteAsync(InkCanvas canvas, string pageTitle)
    {
        if (string.IsNullOrEmpty(_accessToken))
        {
            if (!await AuthenticateAsync())
                return false;
        }

        try
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            // 1. 获取笔记本 ID
            var notebookId = await GetNotebookIdAsync(http);
            if (string.IsNullOrEmpty(notebookId))
            {
                MessageBox.Show($"未找到笔记本 '{_config.OneNote.NotebookName}'", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // 2. 获取或创建分区
            var sectionId = await GetOrCreateSectionAsync(http, notebookId);
            if (string.IsNullOrEmpty(sectionId))
            {
                MessageBox.Show("无法获取或创建分区", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // 3. 提取文字内容和渲染笔画图片
            var textContents = ExtractTextContents(canvas);
            var hasStrokes = canvas.Strokes.Count > 0;
            var strokeImageBase64 = hasStrokes ? RenderStrokesToBase64(canvas) : null;

            // 4. 构建 HTML 内容
            var bodyContent = new StringBuilder();
            bodyContent.AppendLine($"    <h1>{EscapeHtml(pageTitle)}</h1>");
            bodyContent.AppendLine($"    <p>创建时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");
            
            // 添加文字内容
            if (textContents.Count > 0)
            {
                bodyContent.AppendLine("    <div>");
                foreach (var text in textContents)
                {
                    bodyContent.AppendLine($"      <p style=\"font-size:{text.FontSize}pt;color:{text.Color};\">{EscapeHtml(text.Text)}</p>");
                }
                bodyContent.AppendLine("    </div>");
            }
            
            // 添加笔画图片
            if (hasStrokes && strokeImageBase64 != null)
            {
                bodyContent.AppendLine($"    <img src=\"data:image/png;base64,{strokeImageBase64}\" alt=\"笔画内容\" />");
            }

            var html = $@"<!DOCTYPE html>
<html>
  <head>
    <title>{EscapeHtml(pageTitle)}</title>
  </head>
  <body>
{bodyContent}
  </body>
</html>";

            var content = new StringContent(html, Encoding.UTF8, "application/xhtml+xml");
            var response = await http.PostAsync(
                $"https://graph.microsoft.com/v1.0/me/onenote/sections/{sectionId}/pages",
                content);

            if (response.IsSuccessStatusCode)
            {
                return true;
            }
            else
            {
                var error = await response.Content.ReadAsStringAsync();
                MessageBox.Show($"保存失败: {error}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存到 OneNote 失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    /// <summary>
    /// HTML 转义
    /// </summary>
    private string EscapeHtml(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>
    /// 提取画布上的文字内容
    /// </summary>
    private List<TextContent> ExtractTextContents(InkCanvas canvas)
    {
        var texts = new List<TextContent>();
        
        foreach (UIElement child in canvas.Children)
        {
            if (child is TextBlock textBlock && !string.IsNullOrWhiteSpace(textBlock.Text))
            {
                var color = "#000000";
                if (textBlock.Foreground is SolidColorBrush brush)
                {
                    color = $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                }
                
                texts.Add(new TextContent
                {
                    Text = textBlock.Text,
                    FontSize = (int)(textBlock.FontSize * 0.75), // 像素转点
                    Color = color,
                    Left = InkCanvas.GetLeft(child),
                    Top = InkCanvas.GetTop(child)
                });
            }
        }
        
        // 按位置排序（从上到下，从左到右）
        texts.Sort((a, b) =>
        {
            var topDiff = a.Top.CompareTo(b.Top);
            return topDiff != 0 ? topDiff : a.Left.CompareTo(b.Left);
        });
        
        return texts;
    }

    /// <summary>
    /// 只渲染笔画为图片（不包含文字）
    /// </summary>
    private string RenderStrokesToBase64(InkCanvas canvas)
    {
        var width = (int)canvas.ActualWidth;
        var height = (int)canvas.ActualHeight;

        if (width <= 0 || height <= 0)
        {
            width = 800;
            height = 600;
        }

        var renderBitmap = new RenderTargetBitmap(
            width, height, 96, 96, PixelFormats.Pbgra32);

        var visual = new DrawingVisual();
        using (var context = visual.RenderOpen())
        {
            // 透明背景
            context.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, width, height));

            // 只绘制笔画
            foreach (var stroke in canvas.Strokes)
            {
                stroke.Draw(context);
            }
        }

        renderBitmap.Render(visual);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

        using var stream = new MemoryStream();
        encoder.Save(stream);
        return Convert.ToBase64String(stream.ToArray());
    }

    private async Task<string?> GetNotebookIdAsync(HttpClient http)
    {
        var response = await http.GetAsync("https://graph.microsoft.com/v1.0/me/onenote/notebooks");
        var json = await response.Content.ReadAsStringAsync();
        var doc = JsonDocument.Parse(json);

        foreach (var notebook in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            if (notebook.GetProperty("displayName").GetString() == _config.OneNote.NotebookName)
            {
                return notebook.GetProperty("id").GetString();
            }
        }

        return null;
    }

    private async Task<string?> GetOrCreateSectionAsync(HttpClient http, string notebookId)
    {
        // 查找现有分区
        var response = await http.GetAsync(
            $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sections");
        var json = await response.Content.ReadAsStringAsync();
        var doc = JsonDocument.Parse(json);

        foreach (var section in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            if (section.GetProperty("displayName").GetString() == _config.OneNote.SectionName)
            {
                return section.GetProperty("id").GetString();
            }
        }

        // 创建新分区
        var newSectionJson = JsonSerializer.Serialize(new { displayName = _config.OneNote.SectionName });
        var newSectionContent = new StringContent(newSectionJson, Encoding.UTF8, "application/json");
        var createResponse = await http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebookId}/sections",
            newSectionContent);

        if (createResponse.IsSuccessStatusCode)
        {
            var newSectionJsonResult = await createResponse.Content.ReadAsStringAsync();
            var newSectionDoc = JsonDocument.Parse(newSectionJsonResult);
            return newSectionDoc.RootElement.GetProperty("id").GetString();
        }

        return null;
    }

}

/// <summary>
/// 文字内容
/// </summary>
public class TextContent
{
    public string Text { get; set; } = string.Empty;
    public int FontSize { get; set; } = 12;
    public string Color { get; set; } = "#000000";
    public double Left { get; set; }
    public double Top { get; set; }
}
