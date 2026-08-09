using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;

namespace GBT9704_2012排版工具;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        // 注册全局异常兜底：UI 线程异常与后台任务未观察异常都落盘日志，避免进程无声崩溃
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException += App_UnobservedTaskException;
    }

    private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
    {
        落盘异常日志("UI线程未处理异常", e.Exception);
        System.Windows.MessageBox.Show(
            $"发生未处理的异常：{e.Exception.Message}\r\n\r\n详细信息已写入错误日志。",
            "错误",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error);
        e.Handled = true;
    }

    private void App_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        落盘异常日志("后台任务未观察异常", e.Exception);
        e.SetObserved();
    }

    private static void 落盘异常日志(string 来源, Exception ex)
    {
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB-T9704-2012排版工具");
            Directory.CreateDirectory(dir);
            File.AppendAllText(
                Path.Combine(dir, "error.log"),
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {来源}{Environment.NewLine}{ex}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
            // 日志落盘失败不再向外抛，避免二次异常
        }
    }
}


