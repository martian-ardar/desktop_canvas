using System.Windows;
using System.Windows.Threading;

namespace TransparentCanvas.Controls;

/// <summary>
/// 日程提醒弹出窗口 - 每隔一秒抖动一秒
/// </summary>
public partial class ReminderWindow : Window
{
    private readonly DispatcherTimer _shakeTimer;
    private readonly DispatcherTimer _shakeAnimTimer;
    private bool _isShaking = false;
    private int _shakeStep = 0;
    private double _originalLeft;
    private double _originalTop;
    
    // 抖动偏移序列（左右上下抖动）
    private static readonly (double dx, double dy)[] ShakeOffsets = new[]
    {
        (8.0, 0.0), (-8.0, 0.0), (6.0, -4.0), (-6.0, 4.0),
        (4.0, 0.0), (-4.0, 0.0), (3.0, -2.0), (-3.0, 2.0),
        (2.0, 0.0), (-2.0, 0.0), (0.0, 0.0)
    };

    public ReminderWindow(string noteTitle, DateTime targetTime)
    {
        InitializeComponent();
        
        NoteTitleText.Text = noteTitle;
        TimeText.Text = $"提醒时间: {targetTime:yyyy-MM-dd HH:mm}";
        
        // 每秒切换一次：抖动/静止
        _shakeTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _shakeTimer.Tick += ShakeTimer_Tick;
        
        // 抖动动画定时器 (每30ms一帧)
        _shakeAnimTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(30)
        };
        _shakeAnimTimer.Tick += ShakeAnimTimer_Tick;
        
        this.Loaded += ReminderWindow_Loaded;
    }
    
    private void ReminderWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _originalLeft = this.Left;
        _originalTop = this.Top;
        _shakeTimer.Start();
        
        // 播放系统提示音
        System.Media.SystemSounds.Exclamation.Play();
    }
    
    /// <summary>
    /// 每秒切换抖动/静止状态
    /// </summary>
    private void ShakeTimer_Tick(object? sender, EventArgs e)
    {
        _isShaking = !_isShaking;
        
        if (_isShaking)
        {
            _shakeStep = 0;
            _originalLeft = this.Left;
            _originalTop = this.Top;
            _shakeAnimTimer.Start();
        }
        else
        {
            _shakeAnimTimer.Stop();
            // 恢复原位
            this.Left = _originalLeft;
            this.Top = _originalTop;
        }
    }
    
    /// <summary>
    /// 抖动动画帧
    /// </summary>
    private void ShakeAnimTimer_Tick(object? sender, EventArgs e)
    {
        if (_shakeStep < ShakeOffsets.Length)
        {
            var offset = ShakeOffsets[_shakeStep];
            this.Left = _originalLeft + offset.dx;
            this.Top = _originalTop + offset.dy;
            _shakeStep++;
        }
        else
        {
            // 一轮抖动完成后重新开始
            _shakeStep = 0;
        }
    }
    
    /// <summary>
    /// 关闭按钮点击
    /// </summary>
    private void DismissButton_Click(object sender, RoutedEventArgs e)
    {
        _shakeTimer.Stop();
        _shakeAnimTimer.Stop();
        this.Close();
    }
    
    protected override void OnClosed(EventArgs e)
    {
        _shakeTimer.Stop();
        _shakeAnimTimer.Stop();
        base.OnClosed(e);
    }
}