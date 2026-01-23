using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;

namespace TransparentCanvas.Models;

/// <summary>
/// 页面数据模型 - 存储单个页面的状态
/// </summary>
public class PageData
{
    private static int _nextId = 1;

    public int Id { get; }
    public string Title { get; set; }
    public StrokeCollection Strokes { get; set; }
    public List<ChildElementData> Children { get; set; }
    public DateTime CreatedAt { get; }
    public DateTime ModifiedAt { get; set; }

    public PageData(string? title = null)
    {
        Id = _nextId++;
        Title = title ?? $"页面 {Id}";
        Strokes = new StrokeCollection();
        Children = new List<ChildElementData>();
        CreatedAt = DateTime.Now;
        ModifiedAt = DateTime.Now;
    }

    /// <summary>
    /// 从 InkCanvas 保存状态
    /// </summary>
    public void SaveFromCanvas(InkCanvas canvas)
    {
        // 保存笔画
        Strokes = canvas.Strokes.Clone();

        // 保存子元素
        Children.Clear();
        foreach (UIElement child in canvas.Children)
        {
            var childData = new ChildElementData
            {
                Element = child,
                Left = InkCanvas.GetLeft(child),
                Top = InkCanvas.GetTop(child)
            };
            Children.Add(childData);
        }

        ModifiedAt = DateTime.Now;
    }

    /// <summary>
    /// 恢复状态到 InkCanvas
    /// </summary>
    public void RestoreToCanvas(InkCanvas canvas)
    {
        // 清除当前内容
        canvas.Strokes.Clear();
        canvas.Children.Clear();

        // 恢复笔画
        foreach (var stroke in Strokes)
        {
            canvas.Strokes.Add(stroke.Clone());
        }

        // 恢复子元素
        foreach (var childData in Children)
        {
            canvas.Children.Add(childData.Element);
            InkCanvas.SetLeft(childData.Element, childData.Left);
            InkCanvas.SetTop(childData.Element, childData.Top);
        }
    }
}

/// <summary>
/// 子元素数据 - 存储 InkCanvas 子元素的位置信息
/// </summary>
public class ChildElementData
{
    public UIElement Element { get; set; } = null!;
    public double Left { get; set; }
    public double Top { get; set; }
}
