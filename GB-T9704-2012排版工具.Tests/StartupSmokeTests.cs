using System.Threading;

namespace GB_T9704_2012排版工具.Tests;

public sealed class 启动冒烟测试
{
    [Fact]
    public void 主窗口必须能够创建()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var window = new GBT9704_2012排版工具.MainWindow();
                var modes = window.FindName("_modeBox") as System.Windows.Controls.ComboBox;
                if (modes?.Items.Count != 3) throw new InvalidOperationException("主窗口缺少三种排版选择");
                window.Close();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        Assert.Null(failure);
    }
}
