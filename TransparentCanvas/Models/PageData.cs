using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;

namespace TransparentCanvas.Models;

/// <summary>
/// 笔记类型
/// </summary>
public enum NoteType
{
    /// <summary>
    /// 普通笔记
    /// </summary>
    Normal,

    /// <summary>
    /// 日程提醒类笔记
    /// </summary>
    ScheduleReminder
}

/// <summary>
/// 页面数据模型 - 存储单个页面的状态
/// </summary>
public class PageData
{
    private static int _nextId = 1;

    /// <summary>
    /// 预览下一个 ID（用于对话框默认标题）
    /// </summary>
    public static int _nextIdPreview => _nextId;

    public int Id { get; }
    public string Title { get; set; }
    public NoteType NoteType { get; set; }
    public StrokeCollection Strokes { get; set; }
    public List<ChildElementData> Children { get; set; }
    public DateTime CreatedAt { get; }
    public DateTime ModifiedAt { get; set; }

    /// <summary>
    /// 日程提醒的目标时间（仅日程提醒类笔记有效）
    /// </summary>
    public DateTime? TargetTime { get; set; }

    /// <summary>
    /// 是否已经触发过提醒
    /// </summary>
    public bool HasReminded { get; set; }

    /// <summary>
    /// 本地存储的唯一标识（GUID），用于文件夹名
    /// </summary>
    public string StorageId { get; set; }

    public PageData(string? title = null, NoteType noteType = NoteType.Normal)
    {
        Id = _nextId++;
        NoteType = noteType;
        Title = title ?? (noteType == NoteType.ScheduleReminder ? $"提醒 {Id}" : $"页面 {Id}");
        Strokes = new StrokeCollection();
        Children = new List<ChildElementData>();
        CreatedAt = DateTime.Now;
        ModifiedAt = DateTime.Now;
        StorageId = Guid.NewGuid().ToString("N");
    }

    /// <summary>
    /// 从本地存储加载时使用的构造函数
    /// </summary>
    public PageData(string storageId, string title, NoteType noteType, DateTime createdAt,
        DateTime modifiedAt, DateTime? targetTime, bool hasReminded)
    {
        Id = _nextId++;
        StorageId = storageId;
        Title = title;
        NoteType = noteType;
        CreatedAt = createdAt;
        ModifiedAt = modifiedAt;
        TargetTime = targetTime;
        HasReminded = hasReminded;
        Strokes = new StrokeCollection();
        Children = new List<ChildElementData>();
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