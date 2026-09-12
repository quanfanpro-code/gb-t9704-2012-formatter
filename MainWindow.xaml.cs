using System.IO;
using System.Windows.Media;
using System.Windows.Documents;
using System.ComponentModel;
using GBT9704_2012排版工具.Contracts;
using GBT9704_2012排版工具.Core;
using Wpf.Ui.Controls;

namespace GBT9704_2012排版工具;

public partial class MainWindow : FluentWindow
{
    private bool _处理中;
    private bool _关闭请求中;
    private CancellationTokenSource? _cts;
    private readonly GovAuditService _auditService = new();
    private FlowDocument _logDoc = null!;
    private 排版模式 _本次模式;
    private double _本次套打高度;

    private static readonly SolidColorBrush _最小化蓝 = new(System.Windows.Media.Color.FromArgb(0xFF, 0x1A, 0x73, 0xE8));
    private static readonly SolidColorBrush _最大化绿 = new(System.Windows.Media.Color.FromArgb(0xFF, 0x10, 0x7C, 0x10));
    private static readonly SolidColorBrush _关闭红 = new(System.Windows.Media.Color.FromArgb(0xFF, 0xE8, 0x11, 0x23));

    private static readonly SolidColorBrush _成功色;
    private static readonly SolidColorBrush _警告色;
    private static readonly SolidColorBrush _错误色;
    private static readonly SolidColorBrush _信息色;
    private static readonly SolidColorBrush _时间色;
    private static readonly System.Windows.Media.FontFamily _微软雅黑 = new("Microsoft YaHei");

    static MainWindow()
    {
        _最小化蓝.Freeze(); _最大化绿.Freeze(); _关闭红.Freeze();
        _成功色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(87, 166, 74)); _成功色.Freeze();
        _警告色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 165, 32)); _警告色.Freeze();
        _错误色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 72, 86)); _错误色.Freeze();
        _信息色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(204, 204, 204)); _信息色.Freeze();
        _时间色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(128, 128, 128)); _时间色.Freeze();
    }

    public MainWindow()
    {
        InitializeComponent();
        _logDoc = _logBox.Document;

        _startButton.Content = "开始处理";
        _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;

        _logDoc.Blocks.Clear();
        AppendLog("系统就绪，准备处理 GB/T9704—2012 公文文件。", LogType.Info);

        Loaded += (_, _) =>
        {
            foreach (var btn in FindVisualChildren<TitleBarButton>(this))
            {
                var capturedBtn = btn;
                var capturedColor = btn.ButtonType switch
                {
                    TitleBarButtonType.Minimize => _最小化蓝,
                    TitleBarButtonType.Maximize => _最大化绿,
                    TitleBarButtonType.Restore => _最大化绿,
                    TitleBarButtonType.Close => _关闭红,
                    _ => null
                };
                if (capturedColor == null) continue;

                DependencyPropertyDescriptor
                    .FromProperty(BackgroundProperty, typeof(TitleBarButton))
                    .AddValueChanged(capturedBtn, (_, _) =>
                    {
                        var bg = capturedBtn.Background as SolidColorBrush;
                        if (bg != null && bg.Color.A == 0)
                            capturedBtn.Background = capturedColor;
                    });
            }
        };
    }

    private void BrowseFile_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title = "选择 GB/T9704—2012 公文 Word 文件",
            Filter = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*"
        };
        if (dlg.ShowDialog() == true)
        {
            _selectedPathBox.Text = dlg.FileName;
            AppendLog($"已选择文件：{dlg.FileName}");
            _statusText.Text = $"已选择文件：{Path.GetFileName(dlg.FileName)}";
        }
    }

    private void BrowseFolder_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "选择待处理 GB/T9704—2012 公文文件夹"
        };
        if (dlg.ShowDialog() == true)
        {
            _selectedPathBox.Text = dlg.FolderName;
            AppendLog($"已选择文件夹：{dlg.FolderName}");
            _statusText.Text = $"已选择文件夹：{Path.GetFileName(dlg.FolderName)}";
        }
    }

    private void BrowseOutputFolder_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "选择排版结果输出目录"
        };
        if (dlg.ShowDialog() == true)
        {
            _outputDirectoryBox.Text = dlg.FolderName;
            AppendLog($"已选择输出目录：{dlg.FolderName}");
        }
    }

    private void ClearOutputFolder_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        _outputDirectoryBox.Text = string.Empty;
        AppendLog("已清空输出目录，后续将输出到原文件所在目录。");
    }

    private void Start_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_处理中)
        {
            try { _cts?.Cancel(); } catch (ObjectDisposedException) { }
            return;
        }

        var path = _selectedPathBox.Text.Trim();
        if (string.IsNullOrEmpty(path))
        {
            System.Windows.MessageBox.Show("请先通过浏览按钮选择文件或文件夹。", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }

        if (!File.Exists(path) && !Directory.Exists(path))
        {
            System.Windows.MessageBox.Show($"路径不存在：{path}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return;
        }

        var 缺失字体 = new GovFontCheckService().获取缺失字体();
        if (缺失字体.Count > 0)
        {
            var answer = System.Windows.MessageBox.Show(
                "检测到以下字体缺失，继续处理可能导致排版效果不准确：\n\n" +
                string.Join("、", 缺失字体) +
                "\n\n是否仍然继续？",
                "字体提示",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (answer != System.Windows.MessageBoxResult.Yes)
            {
                AppendLog("用户取消了本次处理，因为系统缺少必需字体。", LogType.Warning);
                _statusText.Text = "已取消";
                return;
            }

            AppendLog($"检测到缺失字体：{string.Join("、", 缺失字体)}，用户选择继续。", LogType.Warning);
        }

        _本次模式 = (排版模式)_modeBox.SelectedIndex;
        _本次套打高度 = 110;
        if (_本次模式 == 排版模式.套打 &&
            (!double.TryParse(_printHeightBox.Text, out _本次套打高度) || !double.IsFinite(_本次套打高度) || _本次套打高度 is < 70 or > 150))
        {
            System.Windows.MessageBox.Show("请填写有效的纸顶至红线距离（70 至 150 毫米）。", "套打设置");
            return;
        }
        _modeBox.IsEnabled = false;
        _printHeightBox.IsEnabled = false;
        _处理中 = true;
        _cts = new CancellationTokenSource();
        _startButton.Content = "取消处理";
        _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Danger;
        _progressBar.Value = 0;
        _progressLabel.Text = "";
        _statusText.Text = "正在处理";
        _logDoc.Blocks.Clear();

        var includeSubfolders = _includeSubfoldersCheck.IsChecked == true;
        var outputDir = 读取输出目录();
        var token = _cts.Token;
        _ = Task.Run(() => 后台处理(path, includeSubfolders, outputDir, token));
    }

    private void 后台处理(string path, bool includeSubfolders, string? outputDir, CancellationToken token)
    {
        try
        {
            if (File.Exists(path))
            {
                if (token.IsCancellationRequested) return;
                Dispatcher.BeginInvoke(() => AppendLog($"开始处理文件：{path}"));
                var result = 处理单个文件(path, outputDir);
                Dispatcher.Invoke(() =>
                {
                    if (result.Success)
                    {
                        AppendLog($"输出文件：{result.OutputPath}", LogType.Success);
                        if (!string.IsNullOrWhiteSpace(result.AuditPath))
                            AppendLog($"审计摘要：{result.AuditPath}");
                        if (!string.IsNullOrWhiteSpace(result.Message))
                            AppendLog($"校验警告：{result.Message}", LogType.Warning);
                        _statusText.Text = $"处理完成：{Path.GetFileName(result.OutputPath)}";
                    }
                    else
                    {
                        AppendLog(构建失败日志(result), LogType.Error);
                        if (!string.IsNullOrWhiteSpace(result.AuditPath))
                            AppendLog($"审计摘要：{result.AuditPath}");
                        _statusText.Text = "处理失败";
                    }
                    _progressBar.Value = 100;
                    _progressLabel.Text = "1/1";
                });
            }
            else
            {
                // 勾选包含子目录时忽略无权限的子目录，避免整批枚举被单个目录异常终止
                var allFiles = includeSubfolders
                    ? Directory.GetFiles(path, "*.docx", new EnumerationOptions
                    {
                        IgnoreInaccessible = true,
                        RecurseSubdirectories = true
                    })
                    : Directory.GetFiles(path, "*.docx", SearchOption.TopDirectoryOnly);
                var files = allFiles
                    .Where(f => !输出文件命名规则.是已排版文件(f))
                    .ToList();
                var 跳过已排版数 = allFiles.Length - files.Count;
                if (跳过已排版数 > 0)
                    Dispatcher.BeginInvoke(() => AppendLog($"跳过 {跳过已排版数} 个已排版文件。"));

                if (files.Count == 0)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AppendLog("未找到需要处理的 docx 文件。");
                        _statusText.Text = "未找到文件";
                    });
                    return;
                }

                Dispatcher.BeginInvoke(() => AppendLog($"开始批量处理：{path}（共 {files.Count} 个文件）"));

                var success = 0;
                var fail = 0;
                var batchAuditItems = new List<BatchAuditItemContract>();
                for (var i = 0; i < files.Count; i++)
                {
                    if (token.IsCancellationRequested) break;
                    var file = files[i];
                    var idx = i + 1;
                    try
                    {
                        var result = 处理单个文件(file, outputDir);
                        batchAuditItems.Add(构建批量审计项(file, result));
                        if (result.Success)
                            success++;
                        else
                            fail++;
                        if (token.IsCancellationRequested) break;
                        // 逐文件日志走异步派发，后台线程不再同步等待 UI
                        Dispatcher.BeginInvoke(() =>
                        {
                            if (result.Success)
                            {
                                AppendLog($"[{idx}/{files.Count}] ✓ {Path.GetFileName(file)} → {Path.GetFileName(result.OutputPath)}", LogType.Success);
                                if (!string.IsNullOrWhiteSpace(result.Message))
                                    AppendLog($"[{idx}/{files.Count}] 校验警告：{result.Message}", LogType.Warning);
                            }
                            else
                            {
                                AppendLog($"[{idx}/{files.Count}] ✗ {Path.GetFileName(file)} - {构建失败摘要(result)}", LogType.Error);
                            }
                            _progressBar.Value = (double)idx / files.Count * 100;
                            _progressLabel.Text = $"{idx}/{files.Count}";
                            _statusText.Text = $"正在处理：{Path.GetFileName(file)}";
                        });
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        fail++;
                        batchAuditItems.Add(new BatchAuditItemContract(
                            file,
                            null,
                            0,
                            false,
                            0,
                            0,
                            "阻断",
                            "文件访问",
                            "FILE_LOCKED",
                            "文件占用",
                            false,
                            ex.Message,
                            null,
                            null));
                        if (!token.IsCancellationRequested)
                        {
                            Dispatcher.BeginInvoke(() =>
                                AppendLog($"[{idx}/{files.Count}] ⚠ 文件被占用：{Path.GetFileName(file)} - {ex.Message}", LogType.Warning));
                        }
                    }
                    catch (IOException ex)
                    {
                        fail++;
                        batchAuditItems.Add(new BatchAuditItemContract(
                            file,
                            null,
                            0,
                            false,
                            0,
                            0,
                            "阻断",
                            "文件复制",
                            "FILE_IO_ERROR",
                            "文件读写异常",
                            true,
                            ex.Message,
                            null,
                            null));
                        if (!token.IsCancellationRequested)
                        {
                            Dispatcher.BeginInvoke(() =>
                                AppendLog($"[{idx}/{files.Count}] ✗ 文件读写异常：{Path.GetFileName(file)} - {ex.Message}", LogType.Error));
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        batchAuditItems.Add(new BatchAuditItemContract(
                            file,
                            null,
                            0,
                            false,
                            0,
                            0,
                            "阻断",
                            "运行异常",
                            "UNHANDLED_EXCEPTION",
                            "未处理异常",
                            true,
                            ex.Message,
                            null,
                            null));
                        if (!token.IsCancellationRequested)
                        {
                            Dispatcher.BeginInvoke(() =>
                                AppendLog($"[{idx}/{files.Count}] ✗ {Path.GetFileName(file)} - {提炼异常日志(ex)}", LogType.Error));
                        }
                    }
                }

                if (batchAuditItems.Count > 0)
                {
                    var batchAuditPath = _auditService.写入批量审计清单(path, batchAuditItems, outputDir);
                    Dispatcher.BeginInvoke(() => AppendLog($"批量审计清单：{batchAuditPath}"));
                }

                if (token.IsCancellationRequested)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AppendLog($"⚠ 用户已取消处理，已完成 {success + fail}/{files.Count} 个文件", LogType.Warning);
                        _statusText.Text = $"已取消：完成 {success + fail}/{files.Count}";
                    });
                }
                else
                {
                    Dispatcher.Invoke(() =>
                        _statusText.Text = $"批量处理完成：成功 {success}，失败 {fail}");
                }
            }
        }
        catch (Exception ex)
        {
            if (!token.IsCancellationRequested)
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"【严重错误】{提炼异常日志(ex)}", LogType.Error);
                    _statusText.Text = "处理失败";
                });
            }
        }
        finally
        {
            Dispatcher.Invoke(() =>
            {
                _处理中 = false;
                _modeBox.IsEnabled = true;
                _printHeightBox.IsEnabled = true;
                _startButton.Content = "开始处理";
                _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
                if (token.IsCancellationRequested)
                {
                    AppendLog("处理已被取消。");
                    _statusText.Text = "已取消";
                }
            });
            var oldCts = Interlocked.Exchange(ref _cts, null);
            oldCts?.Dispose();
            // 处理中用户请求过关窗：任务已收尾，现在真正执行关窗
            if (_关闭请求中)
            {
                _关闭请求中 = false;
                Dispatcher.BeginInvoke(Close);
            }
        }
    }

    private ResponseContract 处理单个文件(string inputPath, string? outputDir)
    {
        var request = new RequestContract
        {
            InputPath = inputPath,
            OutputPath = 输出文件命名规则.生成输出路径(inputPath, outputDir),
            模式 = _本次模式,
            套打红线距纸顶毫米 = _本次套打高度
        };

        if (_本次模式 != 排版模式.普通材料)
        {
            using var source = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(inputPath, false);
            var main = source.MainDocumentPart ?? throw new InvalidOperationException("文档缺少正文。");
            request.要素 = 公文要素服务.读取(main);
            var 原文 = string.Join(Environment.NewLine, main.Document!.Body!.Elements().Select(p => p.InnerText));
            var accepted = Dispatcher.Invoke(() =>
            {
                if (_cts?.IsCancellationRequested == true) return false;
                var dialog = new 公文要素窗口(inputPath, 原文, request.要素) { Owner = this };
                return dialog.ShowDialog() == true;
            });
            if (!accepted) return ResponseContract.Fail("已取消本文件的要素确认，原稿未更改。", errorCode: "USER_CANCELLED");
        }

        var pipeline = new GovDocumentPipeline();
        var result = pipeline.Process(request);

        Dispatcher.BeginInvoke(() =>
        {
            if (!string.IsNullOrWhiteSpace(result.DocumentKind))
            {
                AppendLog($"识别文种：{result.DocumentKind}；依据：{result.DocumentKindReason ?? "无"}");
            }

            AppendLog(
                $"结构摘要：附件 {result.AttachmentCount}，页码字段 {(result.HasPageNumberField ? "有" : "无")}，横向节 {result.LandscapeSectionCount}，版记 {result.ImprintCount}");
        });

        return result;
    }

    private static string 构建失败日志(ResponseContract result)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(result.FailureStage))
            parts.Add($"阶段：{result.FailureStage}");
        if (!string.IsNullOrWhiteSpace(result.RuleCode))
            parts.Add($"规则编号：{result.RuleCode}");
        if (!string.IsNullOrWhiteSpace(result.RuleName))
            parts.Add($"规则名称：{result.RuleName}");
        if (result.NeedsManualReview)
            parts.Add("需人工复核：是");

        parts.Add($"原因：{result.Message ?? result.ErrorCode ?? "未知错误"}");
        return $"流水线返回失败：{string.Join("；", parts)}";
    }

    private void Mode_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_printHeightPanel is not null)
            _printHeightPanel.Visibility = _modeBox.SelectedIndex == 2 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
    }

    private static string 构建失败摘要(ResponseContract result)
    {
        return string.Join("；", new[]
        {
            result.FailureStage,
            result.RuleCode,
            result.Message ?? result.ErrorCode
        }.Where(x => !string.IsNullOrWhiteSpace(x)));
    }

    private static BatchAuditItemContract 构建批量审计项(string inputPath, ResponseContract result)
    {
        return new BatchAuditItemContract(
            inputPath,
            string.IsNullOrWhiteSpace(result.OutputPath) ? null : result.OutputPath,
            result.AttachmentCount,
            result.HasPageNumberField,
            result.LandscapeSectionCount,
            result.ImprintCount,
            result.Success ? (result.NeedsManualReview ? "人工复核" : "通过") : "阻断",
            result.FailureStage,
            result.RuleCode,
            result.RuleName,
            result.NeedsManualReview,
            result.Message,
            result.DocumentKind,
            result.AuditPath);
    }

    private static string 提炼异常日志(Exception ex)
    {
        var text = (ex.Message ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
        return string.IsNullOrEmpty(text) ? "文档结构异常或损坏" : text;
    }

    public enum LogType { Info, Success, Warning, Error }

    private string? 读取输出目录()
    {
        var text = _outputDirectoryBox?.Text?.Trim();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private void AppendLog(string message, LogType type = LogType.Info)
    {
        if (_logDoc == null || _logBox == null)
            return;

        var timeStr = DateTime.Now.ToString("HH:mm:ss");
        string prefix = $"> [{timeStr}] ";

        // 级别由调用方显式传入，只决定图标和颜色，不嗅探内容、不改动文本
        var (icon, colorBrush) = type switch
        {
            LogType.Success => ("🟢 ", _成功色),
            LogType.Warning => ("🟡 ", _警告色),
            LogType.Error => ("🔴 ", _错误色),
            _ => ("", _信息色),
        };

        var paragraph = new Paragraph();
        paragraph.Inlines.Add(new Run(prefix) { Foreground = _时间色, FontFamily = _微软雅黑 });
        paragraph.Inlines.Add(new Run(icon + message) { Foreground = colorBrush, FontFamily = _微软雅黑 });

        _logDoc.Blocks.Add(paragraph);
        _logBox.ScrollToEnd();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_处理中)
        {
            // 后台任务可能正在写 docx，直接放行关窗会留下半截文件：
            // 先阻止关窗并请求取消，等任务在 finally 收尾后再由它发起真正的关窗
            e.Cancel = true;
            _关闭请求中 = true;
            AppendLog("正在等待后台任务收尾，完成后窗口将自动关闭……", LogType.Warning);
            try { _cts?.Cancel(); } catch (ObjectDisposedException) { }
            return;
        }
        base.OnClosing(e);
    }

    private static IEnumerable<T> FindVisualChildren<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
    {
        if (parent == null) yield break;
        var children = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < children; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T typed) yield return typed;
            foreach (var desc in FindVisualChildren<T>(child)) yield return desc;
        }
    }
}
