using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using TransparentCanvas.Models;
using TransparentCanvas.Services;

namespace TransparentCanvas;

/// <summary>
/// 透明画板主窗口
/// </summary>
public partial class MainWindow : Window
{
    private Color _currentColor = Colors.Black;
    private double _currentThickness = 3;
    
    // 页面管理
    private readonly List<PageData> _pages = new();
    private PageData? _currentPage;
    
    // OneNote 服务
    private readonly OneNoteService _oneNoteService = new();
    
    // 撤销历史
    private readonly Stack<UndoAction> _undoStack = new();
    private bool _isUndoing = false;

    public MainWindow()
    {
        InitializeComponent();
        InitializeInkCanvas();
        InitializePages();
        CreateAppIcon();
        
        // 注册键盘快捷键 (使用 PreviewKeyDown 防止 InkCanvas 拦截)
        this.PreviewKeyDown += MainWindow_KeyDown;
        
        // 注册粘贴命令
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Paste, OnPaste));
        
        // 应用圆角裁剪
        this.Loaded += MainWindow_Loaded;
        this.SizeChanged += MainWindow_SizeChanged;
    }

    /// <summary>
    /// 创建应用图标
    /// </summary>
    private void CreateAppIcon()
    {
        // 创建一个简单的画笔图标
        var size = 64;
        var visual = new DrawingVisual();
        using (var context = visual.RenderOpen())
        {
            // 背景圆
            context.DrawEllipse(
                new SolidColorBrush(Color.FromRgb(45, 45, 48)),
                new Pen(new SolidColorBrush(Color.FromRgb(80, 80, 80)), 1),
                new Point(size / 2, size / 2),
                size / 2 - 2, size / 2 - 2);

            // 画笔形状
            var penGeometry = new StreamGeometry();
            using (var ctx = penGeometry.Open())
            {
                ctx.BeginFigure(new Point(45, 15), true, true);
                ctx.LineTo(new Point(52, 22), true, false);
                ctx.LineTo(new Point(25, 49), true, false);
                ctx.LineTo(new Point(15, 52), true, false);
                ctx.LineTo(new Point(18, 42), true, false);
            }
            context.DrawGeometry(
                new SolidColorBrush(Color.FromRgb(0, 120, 212)),
                null,
                penGeometry);

            // 笔尖
            var tipGeometry = new StreamGeometry();
            using (var ctx = tipGeometry.Open())
            {
                ctx.BeginFigure(new Point(15, 52), true, true);
                ctx.LineTo(new Point(18, 42), true, false);
                ctx.LineTo(new Point(12, 56), true, false);
            }
            context.DrawGeometry(
                new SolidColorBrush(Color.FromRgb(255, 180, 80)),
                null,
                tipGeometry);
        }

        var renderBitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
        renderBitmap.Render(visual);

        this.Icon = renderBitmap;
    }
    
    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        ApplyRoundedCorners();
        // 确保默认设置为选择模式
        SetSelectMode();
    }
    
    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyRoundedCorners();
    }
    
    private void ApplyRoundedCorners()
    {
        var radius = 6;
        var clip = new RectangleGeometry(
            new Rect(0, 0, this.ActualWidth, this.ActualHeight),
            radius, radius);
        this.Clip = clip;
    }

    #region 浮动工具栏

    /// <summary>
    /// 鼠标进入主区域 - 显示工具栏
    /// </summary>
    private void MainGrid_MouseEnter(object sender, MouseEventArgs e)
    {
        ShowFloatingToolbar();
    }

    /// <summary>
    /// 鼠标离开主区域 - 隐藏工具栏
    /// </summary>
    private void MainGrid_MouseLeave(object sender, MouseEventArgs e)
    {
        HideFloatingToolbar();
    }

    #endregion

    #region 画布透明度

    /// <summary>
    /// 鼠标进入画布区域 - 显示半透明背景
    /// </summary>
    private void CanvasBorder_MouseEnter(object sender, MouseEventArgs e)
    {
        // 设置60%透明度的白色背景
        var animation = new ColorAnimation
        {
            To = Color.FromArgb(153, 255, 255, 255), // 60% 不透明 = 153/255
            Duration = TimeSpan.FromMilliseconds(150),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        var brush = new SolidColorBrush(Colors.Transparent);
        CanvasBorder.Background = brush;
        brush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
    }

    /// <summary>
    /// 鼠标离开画布区域 - 恢复透明
    /// </summary>
    private void CanvasBorder_MouseLeave(object sender, MouseEventArgs e)
    {
        var animation = new ColorAnimation
        {
            To = Colors.Transparent,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };
        if (CanvasBorder.Background is SolidColorBrush brush)
        {
            brush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }
    }

    /// <summary>
    /// 显示浮动工具栏
    /// </summary>
    private void ShowFloatingToolbar()
    {
        var animation = new DoubleAnimation
        {
            To = 1,
            Duration = TimeSpan.FromMilliseconds(150),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        FloatingToolbar.BeginAnimation(OpacityProperty, animation);
    }

    /// <summary>
    /// 隐藏浮动工具栏
    /// </summary>
    private void HideFloatingToolbar()
    {
        var animation = new DoubleAnimation
        {
            To = 0,
            Duration = TimeSpan.FromMilliseconds(300),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };
        FloatingToolbar.BeginAnimation(OpacityProperty, animation);
    }

    #endregion

    /// <summary>
    /// 初始化画布设置
    /// </summary>
    private void InitializeInkCanvas()
    {
        var drawingAttributes = new DrawingAttributes
        {
            Color = _currentColor,
            Width = _currentThickness,
            Height = _currentThickness,
            FitToCurve = true,
            StylusTip = StylusTip.Ellipse
        };
        inkCanvas.DefaultDrawingAttributes = drawingAttributes;
        
        // 监听笔画变化以支持撤销
        inkCanvas.Strokes.StrokesChanged += Strokes_StrokesChanged;
    }
    
    /// <summary>
    /// 笔画变化事件 - 记录撤销历史
    /// </summary>
    private void Strokes_StrokesChanged(object? sender, StrokeCollectionChangedEventArgs e)
    {
        if (_isUndoing) return;
        
        // 记录添加的笔画
        if (e.Added.Count > 0)
        {
            foreach (var stroke in e.Added)
            {
                _undoStack.Push(new UndoAction(UndoActionType.AddStroke, stroke));
            }
        }
        
        // 记录删除的笔画
        if (e.Removed.Count > 0)
        {
            foreach (var stroke in e.Removed)
            {
                _undoStack.Push(new UndoAction(UndoActionType.RemoveStroke, stroke));
            }
        }
    }

    #region 页面管理

    /// <summary>
    /// 初始化页面系统
    /// </summary>
    private void InitializePages()
    {
        // 创建第一个页面
        CreateNewPage();
    }

    /// <summary>
    /// 创建新页面
    /// </summary>
    private void CreateNewPage()
    {
        // 保存当前页面状态
        SaveCurrentPageState();

        // 创建新页面
        var newPage = new PageData();
        _pages.Add(newPage);
        
        // 切换到新页面
        SwitchToPage(newPage);
        
        // 刷新标签栏
        RefreshPageTabs();
    }

    /// <summary>
    /// 保存当前页面状态
    /// </summary>
    private void SaveCurrentPageState()
    {
        if (_currentPage != null)
        {
            _currentPage.SaveFromCanvas(inkCanvas);
        }
    }

    /// <summary>
    /// 切换到指定页面
    /// </summary>
    private void SwitchToPage(PageData page)
    {
        // 保存当前页面
        SaveCurrentPageState();

        // 设置新的当前页面
        _currentPage = page;

        // 恢复页面内容
        page.RestoreToCanvas(inkCanvas);

        // 刷新标签栏高亮
        RefreshPageTabs();
    }

    /// <summary>
    /// 关闭指定页面
    /// </summary>
    private void ClosePage(PageData page)
    {
        // 至少保留一个页面
        if (_pages.Count <= 1)
        {
            MessageBox.Show("至少需要保留一个页面", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var pageIndex = _pages.IndexOf(page);
        _pages.Remove(page);

        // 如果关闭的是当前页面，切换到相邻页面
        if (_currentPage == page)
        {
            var newIndex = Math.Min(pageIndex, _pages.Count - 1);
            _currentPage = null; // 避免保存已删除的页面
            SwitchToPage(_pages[newIndex]);
        }
        else
        {
            RefreshPageTabs();
        }
    }

    /// <summary>
    /// 刷新页面标签栏
    /// </summary>
    private void RefreshPageTabs()
    {
        PageTabsPanel.Children.Clear();

        foreach (var page in _pages)
        {
            var isActive = page == _currentPage;
            var tab = CreatePageTab(page, isActive);
            PageTabsPanel.Children.Add(tab);
        }
    }

    /// <summary>
    /// 创建页面标签
    /// </summary>
    private Border CreatePageTab(PageData page, bool isActive)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(isActive ? Color.FromRgb(0x3C, 0x3C, 0x3C) : Color.FromRgb(0x2D, 0x2D, 0x30)),
            CornerRadius = new CornerRadius(4, 4, 0, 0),
            Padding = new Thickness(8, 4, 4, 4),
            Margin = new Thickness(2, 0, 0, 0),
            Cursor = Cursors.Hand,
            Tag = page
        };

        var stackPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal
        };

        // 页面标题
        var titleText = new TextBlock
        {
            Text = page.Title,
            Foreground = new SolidColorBrush(isActive ? Colors.White : Color.FromRgb(0xAA, 0xAA, 0xAA)),
            VerticalAlignment = VerticalAlignment.Center,
            FontSize = 11
        };
        stackPanel.Children.Add(titleText);

        // 关闭按钮
        var closeButton = new Button
        {
            Content = "✕",
            Width = 16,
            Height = 16,
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Foreground = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)),
            FontSize = 9,
            Cursor = Cursors.Hand,
            Margin = new Thickness(6, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Tag = page
        };
        closeButton.Click += CloseTabButton_Click;
        stackPanel.Children.Add(closeButton);

        border.Child = stackPanel;
        border.MouseLeftButtonDown += PageTab_Click;

        return border;
    }

    /// <summary>
    /// 页面标签点击事件
    /// </summary>
    private void PageTab_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border border && border.Tag is PageData page)
        {
            // 双击重命名
            if (e.ClickCount == 2)
            {
                RenamePageDialog(page);
                e.Handled = true;
                return;
            }
            
            // 单击切换页面
            if (page != _currentPage)
            {
                SwitchToPage(page);
            }
        }
    }

    /// <summary>
    /// 显示重命名对话框
    /// </summary>
    private void RenamePageDialog(PageData page)
    {
        var dialog = new Window
        {
            Title = "重命名页面",
            Width = 350,
            Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            WindowStyle = WindowStyle.ToolWindow
        };

        var grid = new Grid { Margin = new Thickness(20) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var label = new TextBlock
        {
            Text = "请输入新的页面名称：",
            Margin = new Thickness(0, 0, 0, 10)
        };
        Grid.SetRow(label, 0);
        grid.Children.Add(label);

        var textBox = new TextBox
        {
            Text = page.Title,
            Padding = new Thickness(5),
            Margin = new Thickness(0, 0, 0, 15)
        };
        textBox.SelectAll();
        Grid.SetRow(textBox, 1);
        grid.Children.Add(textBox);

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        
        var okButton = new Button
        {
            Content = "确定",
            Width = 70,
            Height = 28,
            Margin = new Thickness(0, 0, 10, 0),
            IsDefault = true
        };
        okButton.Click += (s, e) =>
        {
            if (!string.IsNullOrWhiteSpace(textBox.Text))
            {
                page.Title = textBox.Text.Trim();
                RefreshPageTabs();
                dialog.DialogResult = true;
            }
        };
        buttonPanel.Children.Add(okButton);

        var cancelButton = new Button
        {
            Content = "取消",
            Width = 70,
            Height = 28,
            IsCancel = true
        };
        buttonPanel.Children.Add(cancelButton);

        Grid.SetRow(buttonPanel, 2);
        grid.Children.Add(buttonPanel);

        dialog.Content = grid;
        dialog.ShowDialog();
    }

    /// <summary>
    /// 关闭标签按钮点击事件
    /// </summary>
    private void CloseTabButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true; // 阻止事件冒泡
        if (sender is Button button && button.Tag is PageData page)
        {
            ClosePage(page);
        }
    }

    /// <summary>
    /// 新建页面按钮点击
    /// </summary>
    private void NewPageButton_Click(object sender, RoutedEventArgs e)
    {
        CreateNewPage();
    }

    #endregion

    #region OneNote 集成

    /// <summary>
    /// 保存到 OneNote 按钮点击
    /// </summary>
    private async void SaveToOneNoteButton_Click(object sender, RoutedEventArgs e)
    {
        await SaveCurrentPageToOneNote();
    }

    /// <summary>
    /// 将当前页面保存到 OneNote
    /// </summary>
    private async Task SaveCurrentPageToOneNote()
    {
        if (_currentPage == null) return;

        // 禁用按钮防止重复点击
        SaveToOneNoteButton.IsEnabled = false;
        SaveToOneNoteButton.Content = "⏳";

        try
        {
            var pageTitle = $"{_currentPage.Title}_{DateTime.Now:yyyyMMdd_HHmmss}";
            var success = await _oneNoteService.SaveCanvasToOneNoteAsync(inkCanvas, pageTitle);

            if (success)
            {
                MessageBox.Show($"已成功保存到 OneNote:\n{pageTitle}", "保存成功", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        finally
        {
            SaveToOneNoteButton.IsEnabled = true;
            SaveToOneNoteButton.Content = "💾";
        }
    }

    #endregion

    #region 标题栏操作

    /// <summary>
    /// 标题栏拖拽移动窗口
    /// </summary>
    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            // 双击切换最大化/还原
            if (WindowState == WindowState.Maximized)
                WindowState = WindowState.Normal;
            else
                WindowState = WindowState.Maximized;
        }
        else
        {
            DragMove();
        }
    }

    /// <summary>
    /// 置顶按钮点击
    /// </summary>
    private void PinButton_Click(object sender, RoutedEventArgs e)
    {
        Topmost = PinButton.IsChecked == true;
    }

    /// <summary>
    /// 最小化按钮点击
    /// </summary>
    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    /// <summary>
    /// 关闭按钮点击
    /// </summary>
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    #endregion

    #region 工具模式切换

    /// <summary>
    /// 切换到绘画模式
    /// </summary>
    private void DrawModeButton_Click(object sender, RoutedEventArgs e)
    {
        SetDrawingMode();
    }

    /// <summary>
    /// 切换到文字模式
    /// </summary>
    private void TextModeButton_Click(object sender, RoutedEventArgs e)
    {
        SetTextMode();
    }

    /// <summary>
    /// 切换到选择模式
    /// </summary>
    private void SelectModeButton_Click(object sender, RoutedEventArgs e)
    {
        SetSelectMode();
    }

    /// <summary>
    /// 切换到橡皮擦模式
    /// </summary>
    private void EraserModeButton_Click(object sender, RoutedEventArgs e)
    {
        SetEraserMode();
    }

    private void SetDrawingMode()
    {
        inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
        inkCanvas.Cursor = Cursors.Pen;
        UpdateModeButtons(DrawModeButton);
    }

    private void SetTextMode()
    {
        // 在文字模式下，点击画布会创建文本框
        inkCanvas.EditingMode = InkCanvasEditingMode.None;
        inkCanvas.Cursor = Cursors.IBeam;
        UpdateModeButtons(TextModeButton);
        // 添加文字模式事件（必须在UpdateModeButtons之后）
        inkCanvas.MouseLeftButtonDown += InkCanvas_TextMode_Click;
    }

    private void SetSelectMode()
    {
        inkCanvas.EditingMode = InkCanvasEditingMode.Select;
        inkCanvas.Cursor = Cursors.Arrow;
        UpdateModeButtons(SelectModeButton);
    }

    private void SetEraserMode()
    {
        inkCanvas.EditingMode = InkCanvasEditingMode.EraseByStroke;
        inkCanvas.Cursor = Cursors.Cross;
        UpdateModeButtons(EraserModeButton);
    }

    private void UpdateModeButtons(ToggleButton activeButton)
    {
        // 先移除文字模式的事件处理（避免重复添加）
        inkCanvas.MouseLeftButtonDown -= InkCanvas_TextMode_Click;
        
        DrawModeButton.IsChecked = activeButton == DrawModeButton;
        TextModeButton.IsChecked = activeButton == TextModeButton;
        SelectModeButton.IsChecked = activeButton == SelectModeButton;
        EraserModeButton.IsChecked = activeButton == EraserModeButton;
    }

    #endregion

    #region 文字输入

    /// <summary>
    /// 文字模式下点击画布创建文本框或编辑已有文本
    /// </summary>
    private void InkCanvas_TextMode_Click(object sender, MouseButtonEventArgs e)
    {
        if (TextModeButton.IsChecked != true) return;

        var position = e.GetPosition(inkCanvas);
        
        // 检查是否点击了已有的 TextBlock
        var clickedTextBlock = FindTextBlockAtPosition(position);
        if (clickedTextBlock != null)
        {
            // 点击已有文本，进入编辑模式
            ConvertToEditableTextBox(clickedTextBlock);
            e.Handled = true;
        }
        else
        {
            // 点击空白区域，创建新的文本框
            CreateTextBox(position);
        }
    }

    /// <summary>
    /// 在指定位置查找 TextBlock
    /// </summary>
    private TextBlock? FindTextBlockAtPosition(Point position)
    {
        foreach (UIElement child in inkCanvas.Children)
        {
            if (child is TextBlock textBlock)
            {
                var left = InkCanvas.GetLeft(textBlock);
                var top = InkCanvas.GetTop(textBlock);
                var width = textBlock.ActualWidth > 0 ? textBlock.ActualWidth : 100;
                var height = textBlock.ActualHeight > 0 ? textBlock.ActualHeight : 30;
                
                var bounds = new Rect(left, top, width, height);
                if (bounds.Contains(position))
                {
                    return textBlock;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// 在指定位置创建文本框
    /// </summary>
    private void CreateTextBox(Point position)
    {
        var textBox = new TextBox
        {
            Background = new SolidColorBrush(Color.FromArgb(200, 255, 255, 255)),
            BorderBrush = new SolidColorBrush(_currentColor),
            BorderThickness = new Thickness(1),
            Foreground = new SolidColorBrush(_currentColor),
            FontSize = _currentThickness * 5 + 10,
            MinWidth = 100,
            MinHeight = 30,
            Padding = new Thickness(4),
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap
        };

        // 失去焦点时转换为固定文本
        textBox.LostFocus += (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                inkCanvas.Children.Remove(textBox);
            }
            else
            {
                var left = InkCanvas.GetLeft(textBox);
                var top = InkCanvas.GetTop(textBox);
                
                // 创建可编辑的 TextBlock
                var textBlock = CreateEditableTextBlock(textBox.Text);
                textBlock.Foreground = textBox.Foreground;
                textBlock.FontSize = textBox.FontSize;
                
                inkCanvas.Children.Remove(textBox);
                inkCanvas.Children.Add(textBlock);
                InkCanvas.SetLeft(textBlock, left);
                InkCanvas.SetTop(textBlock, top);
                
                // 记录到撤销栈
                if (!_isUndoing)
                {
                    _undoStack.Push(new UndoAction(UndoActionType.AddChild, textBlock, left, top));
                }
            }
        };

        inkCanvas.Children.Add(textBox);
        InkCanvas.SetLeft(textBox, position.X);
        InkCanvas.SetTop(textBox, position.Y);
        
        textBox.Focus();
    }

    #endregion

    #region 颜色和粗细

    /// <summary>
    /// 颜色选择变更
    /// </summary>
    private void Color_Checked(object sender, RoutedEventArgs e)
    {
        if (sender is RadioButton radioButton && radioButton.Tag is string colorStr)
        {
            _currentColor = (Color)ColorConverter.ConvertFromString(colorStr);
            UpdateDrawingAttributes();
        }
    }

    /// <summary>
    /// 粗细滑块变更
    /// </summary>
    private void StrokeThicknessSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (ThicknessText != null)
        {
            _currentThickness = e.NewValue;
            ThicknessText.Text = ((int)_currentThickness).ToString();
            UpdateDrawingAttributes();
        }
    }

    /// <summary>
    /// 更新画笔属性
    /// </summary>
    private void UpdateDrawingAttributes()
    {
        if (inkCanvas?.DefaultDrawingAttributes != null)
        {
            inkCanvas.DefaultDrawingAttributes.Color = _currentColor;
            inkCanvas.DefaultDrawingAttributes.Width = _currentThickness;
            inkCanvas.DefaultDrawingAttributes.Height = _currentThickness;
        }
    }

    #endregion

    #region 清除和粘贴

    /// <summary>
    /// 清除画布
    /// </summary>
    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        ClearCanvas();
    }

    private void ClearCanvas()
    {
        _undoStack.Clear(); // 清除时也清空撤销历史
        inkCanvas.Strokes.Clear();
        inkCanvas.Children.Clear();
    }

    /// <summary>
    /// 撤销按钮点击
    /// </summary>
    private void UndoButton_Click(object sender, RoutedEventArgs e)
    {
        Undo();
    }

    /// <summary>
    /// 撤销操作
    /// </summary>
    private void Undo()
    {
        if (_undoStack.Count == 0) return;

        _isUndoing = true;
        try
        {
            var action = _undoStack.Pop();
            
            switch (action.Type)
            {
                case UndoActionType.AddStroke:
                    // 撤销添加笔画 = 移除该笔画
                    if (action.Stroke != null && inkCanvas.Strokes.Contains(action.Stroke))
                    {
                        inkCanvas.Strokes.Remove(action.Stroke);
                    }
                    break;
                    
                case UndoActionType.RemoveStroke:
                    // 撤销移除笔画 = 添加回该笔画
                    if (action.Stroke != null)
                    {
                        inkCanvas.Strokes.Add(action.Stroke);
                    }
                    break;
                    
                case UndoActionType.AddChild:
                    // 撤销添加子元素 = 移除该子元素
                    if (action.Child != null && inkCanvas.Children.Contains(action.Child))
                    {
                        inkCanvas.Children.Remove(action.Child);
                    }
                    break;
                    
                case UndoActionType.RemoveChild:
                    // 撤销移除子元素 = 添加回该子元素
                    if (action.Child != null)
                    {
                        inkCanvas.Children.Add(action.Child);
                        if (action.ChildLeft.HasValue)
                            InkCanvas.SetLeft(action.Child, action.ChildLeft.Value);
                        if (action.ChildTop.HasValue)
                            InkCanvas.SetTop(action.Child, action.ChildTop.Value);
                    }
                    break;
            }
        }
        finally
        {
            _isUndoing = false;
        }
    }

    /// <summary>
    /// 粘贴图片按钮点击
    /// </summary>
    private void PasteImageButton_Click(object sender, RoutedEventArgs e)
    {
        PasteFromClipboard();
    }

    /// <summary>
    /// 粘贴命令处理
    /// </summary>
    private void OnPaste(object sender, ExecutedRoutedEventArgs e)
    {
        PasteFromClipboard();
    }

    /// <summary>
    /// 从剪贴板粘贴内容
    /// </summary>
    private void PasteFromClipboard()
    {
        try
        {
            UIElement? pastedElement = null;
            
            // 尝试获取图片
            if (Clipboard.ContainsImage())
            {
                var image = Clipboard.GetImage();
                if (image != null)
                {
                    pastedElement = CreateImageElement(image);
                }
            }
            // 尝试获取文本
            else if (Clipboard.ContainsText())
            {
                var text = Clipboard.GetText();
                if (!string.IsNullOrEmpty(text))
                {
                    pastedElement = CreateEditableTextBlock(text);
                }
            }
            
            if (pastedElement != null)
            {
                // 计算粘贴位置（画布中心附近）
                var centerX = Math.Max(50, (inkCanvas.ActualWidth - 100) / 2);
                var centerY = Math.Max(50, (inkCanvas.ActualHeight - 100) / 2);
                
                inkCanvas.Children.Add(pastedElement);
                InkCanvas.SetLeft(pastedElement, centerX);
                InkCanvas.SetTop(pastedElement, centerY);
                
                // 记录到撤销栈
                if (!_isUndoing)
                {
                    _undoStack.Push(new UndoAction(UndoActionType.AddChild, pastedElement, centerX, centerY));
                }
                
                // 切换到选择模式并选中粘贴的元素
                SetSelectMode();
                
                // 延迟选中，确保元素已完成布局
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        inkCanvas.Select(new UIElement[] { pastedElement });
                    }
                    catch { /* 忽略选中失败 */ }
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"粘贴失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 创建可编辑的文本块（双击可编辑）
    /// </summary>
    private TextBlock CreateEditableTextBlock(string text)
    {
        var textBlock = new TextBlock
        {
            Text = text,
            Foreground = new SolidColorBrush(_currentColor),
            FontSize = 14,
            Background = Brushes.Transparent,
            Padding = new Thickness(4),
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = 400,
            Cursor = Cursors.Hand
        };
        
        // 双击转换为可编辑的 TextBox
        textBlock.MouseLeftButtonDown += (s, e) =>
        {
            if (e.ClickCount == 2)
            {
                ConvertToEditableTextBox(textBlock);
                e.Handled = true;
            }
        };
        
        return textBlock;
    }

    /// <summary>
    /// 将 TextBlock 转换为可编辑的 TextBox
    /// </summary>
    private void ConvertToEditableTextBox(TextBlock textBlock)
    {
        var left = InkCanvas.GetLeft(textBlock);
        var top = InkCanvas.GetTop(textBlock);
        
        var textBox = new TextBox
        {
            Text = textBlock.Text,
            Foreground = textBlock.Foreground,
            FontSize = textBlock.FontSize,
            Background = new SolidColorBrush(Color.FromArgb(200, 255, 255, 255)),
            BorderBrush = new SolidColorBrush(_currentColor),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4),
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinWidth = 100,
            MaxWidth = 400
        };
        
        // 失去焦点时转换回 TextBlock
        textBox.LostFocus += (s, e) =>
        {
            var newLeft = InkCanvas.GetLeft(textBox);
            var newTop = InkCanvas.GetTop(textBox);
            
            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                inkCanvas.Children.Remove(textBox);
            }
            else
            {
                var newTextBlock = CreateEditableTextBlock(textBox.Text);
                newTextBlock.Foreground = textBox.Foreground;
                newTextBlock.FontSize = textBox.FontSize;
                
                inkCanvas.Children.Remove(textBox);
                inkCanvas.Children.Add(newTextBlock);
                InkCanvas.SetLeft(newTextBlock, newLeft);
                InkCanvas.SetTop(newTextBlock, newTop);
            }
        };
        
        inkCanvas.Children.Remove(textBlock);
        inkCanvas.Children.Add(textBox);
        InkCanvas.SetLeft(textBox, left);
        InkCanvas.SetTop(textBox, top);
        
        textBox.Focus();
        textBox.SelectAll();
    }

    /// <summary>
    /// 创建图片元素
    /// </summary>
    private System.Windows.Controls.Image CreateImageElement(BitmapSource imageSource)
    {
        return new System.Windows.Controls.Image
        {
            Source = imageSource,
            Stretch = Stretch.Uniform,
            MaxWidth = Math.Min(imageSource.PixelWidth, inkCanvas.ActualWidth - 100),
            MaxHeight = Math.Min(imageSource.PixelHeight, inkCanvas.ActualHeight - 100)
        };
    }

    /// <summary>
    /// 添加图片到画布
    /// </summary>
    private void AddImageToCanvas(BitmapSource imageSource)
    {
        var image = CreateImageElement(imageSource);
        inkCanvas.Children.Add(image);
        InkCanvas.SetLeft(image, 50);
        InkCanvas.SetTop(image, 50);
        
        // 切换到选择模式并选中图片
        SetSelectMode();
        
        Dispatcher.BeginInvoke(new Action(() =>
        {
            inkCanvas.Select(new UIElement[] { image });
        }), System.Windows.Threading.DispatcherPriority.Background);
    }

    #endregion

    #region 键盘快捷键

    /// <summary>
    /// 键盘快捷键处理
    /// </summary>
    private async void MainWindow_KeyDown(object sender, KeyEventArgs e)
    {
        // Ctrl+V 粘贴
        if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
        {
            PasteFromClipboard();
            e.Handled = true;
        }
        // Ctrl+Z 撤销
        else if (e.Key == Key.Z && Keyboard.Modifiers == ModifierKeys.Control)
        {
            Undo();
            e.Handled = true;
        }
        // Ctrl+Delete 清除画布
        else if (e.Key == Key.Delete && Keyboard.Modifiers == ModifierKeys.Control)
        {
            ClearCanvas();
            e.Handled = true;
        }
        // Ctrl+N 新建页面
        else if (e.Key == Key.N && Keyboard.Modifiers == ModifierKeys.Control)
        {
            CreateNewPage();
            e.Handled = true;
        }
        // Ctrl+S 保存到 OneNote
        else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
        {
            await SaveCurrentPageToOneNote();
            e.Handled = true;
        }
        // Ctrl+W 关闭当前页面
        else if (e.Key == Key.W && Keyboard.Modifiers == ModifierKeys.Control)
        {
            if (_currentPage != null)
            {
                ClosePage(_currentPage);
            }
            e.Handled = true;
        }
        // D 切换到绘画模式
        else if (e.Key == Key.D && Keyboard.Modifiers == ModifierKeys.None)
        {
            SetDrawingMode();
            e.Handled = true;
        }
        // T 切换到文字模式
        else if (e.Key == Key.T && Keyboard.Modifiers == ModifierKeys.None)
        {
            SetTextMode();
            e.Handled = true;
        }
        // S 切换到选择模式（仅当没有 Ctrl 修饰键时）
        else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.None)
        {
            SetSelectMode();
            e.Handled = true;
        }
        // E 切换到橡皮擦模式
        else if (e.Key == Key.E && Keyboard.Modifiers == ModifierKeys.None)
        {
            SetEraserMode();
            e.Handled = true;
        }
    }

    #endregion

    #region 窗口大小调整

    private void ResizeLeft_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var newWidth = Width - e.HorizontalChange;
        if (newWidth >= MinWidth)
        {
            Width = newWidth;
            Left += e.HorizontalChange;
        }
    }

    private void ResizeRight_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var newWidth = Width + e.HorizontalChange;
        if (newWidth >= MinWidth)
        {
            Width = newWidth;
        }
    }

    private void ResizeTop_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var newHeight = Height - e.VerticalChange;
        if (newHeight >= MinHeight)
        {
            Height = newHeight;
            Top += e.VerticalChange;
        }
    }

    private void ResizeBottom_DragDelta(object sender, DragDeltaEventArgs e)
    {
        var newHeight = Height + e.VerticalChange;
        if (newHeight >= MinHeight)
        {
            Height = newHeight;
        }
    }

    private void ResizeTopLeft_DragDelta(object sender, DragDeltaEventArgs e)
    {
        ResizeLeft_DragDelta(sender, e);
        ResizeTop_DragDelta(sender, e);
    }

    private void ResizeTopRight_DragDelta(object sender, DragDeltaEventArgs e)
    {
        ResizeRight_DragDelta(sender, e);
        ResizeTop_DragDelta(sender, e);
    }

    private void ResizeBottomLeft_DragDelta(object sender, DragDeltaEventArgs e)
    {
        ResizeLeft_DragDelta(sender, e);
        ResizeBottom_DragDelta(sender, e);
    }

    private void ResizeBottomRight_DragDelta(object sender, DragDeltaEventArgs e)
    {
        ResizeRight_DragDelta(sender, e);
        ResizeBottom_DragDelta(sender, e);
    }

    #endregion
}

/// <summary>
/// 撤销操作类型
/// </summary>
public enum UndoActionType
{
    AddStroke,
    RemoveStroke,
    AddChild,
    RemoveChild
}

/// <summary>
/// 撤销操作记录
/// </summary>
public class UndoAction
{
    public UndoActionType Type { get; }
    public Stroke? Stroke { get; }
    public UIElement? Child { get; }
    public double? ChildLeft { get; }
    public double? ChildTop { get; }

    public UndoAction(UndoActionType type, Stroke stroke)
    {
        Type = type;
        Stroke = stroke;
    }

    public UndoAction(UndoActionType type, UIElement child, double? left = null, double? top = null)
    {
        Type = type;
        Child = child;
        ChildLeft = left;
        ChildTop = top;
    }
}
