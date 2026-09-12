using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using GBT9704_2012排版工具.Contracts;
using GBT9704_2012排版工具.Core;

namespace GBT9704_2012排版工具;

public sealed class 公文要素窗口 : Window
{
    public 公文要素窗口(string 文件, string 原文, 公文要素 要素)
    {
        Title = "核对公文要素 · " + Path.GetFileName(文件);
        Width = 960; Height = 760; MinWidth = 800; MinHeight = 600;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        FontFamily = new FontFamily("Microsoft YaHei"); FontSize = 14;
        var root = new DockPanel { Margin = new Thickness(20) };
        Content = root;
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons);
        var cancel = new Button { Content = "取消本文件", IsCancel = true, Padding = new Thickness(16, 8, 16, 8), Margin = new Thickness(0, 0, 12, 0) };
        var save = new Button { Content = "确认并排版", IsDefault = true, Padding = new Thickness(16, 8, 16, 8) };
        buttons.Children.Add(cancel); buttons.Children.Add(save);
        var columns = new Grid(); columns.ColumnDefinitions.Add(new ColumnDefinition()); columns.ColumnDefinitions.Add(new ColumnDefinition()); root.Children.Add(columns);
        var original = new TextBox { Text = 原文, IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 0, 20, 0) };
        columns.Children.Add(original);
        var fields = new StackPanel(); var scroll = new ScrollViewer { Content = fields, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; Grid.SetColumn(scroll, 1); columns.Children.Add(scroll);
        fields.Children.Add(new TextBlock { Text = "请结合左侧原文逐份核对。空白的可选字段不自动补写。", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 12) });
        var boxes = new Dictionary<string, TextBox>();
        foreach (var (property, label) in new[] { ("机关标志", "机关标志（含“文件”等后缀）*"), ("标题", "标题*"), ("文号", "发文字号*"), ("主送机关", "主送机关"), ("署名", "署名"), ("成文日期", "成文日期*"), ("抄送", "抄送（含“抄送：”）"), ("印发机关", "印发机关"), ("印发日期", "印发日期") })
        {
            fields.Children.Add(new TextBlock { Text = label });
            var box = new TextBox { Text = (string?)typeof(公文要素).GetProperty(property)!.GetValue(要素) ?? "", Margin = new Thickness(0, 4, 0, 10), Padding = new Thickness(6), MinHeight = 32 };
            if (property is "标题" or "机关标志" or "主送机关")
            {
                box.AcceptsReturn = true; box.TextWrapping = TextWrapping.Wrap; box.MinHeight = 60;
                box.ToolTip = "长名称或长标题可按词意回车分行，程序保留换行。";
            }
            System.Windows.Automation.AutomationProperties.SetName(box, property);
            boxes.Add(property, box); fields.Children.Add(box);
        }
        var stamp = new CheckBox { Content = "按盖章公文安排署名日期，预留盖章空间", IsChecked = 要素.预留盖章, Margin = new Thickness(0, 6, 0, 12) }; fields.Children.Add(stamp);
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; fields.Children.Add(error);
        save.Click += (_, _) =>
        {
            foreach (var (name, box) in boxes) typeof(公文要素).GetProperty(name)!.SetValue(要素, box.Text.Trim());
            if (要素.机关标志.Length == 0 || 要素.标题.Length == 0 || !公文要素服务.文号有效(要素.文号) || !公文要素服务.日期有效(要素.成文日期))
            {
                error.Text = "请补齐机关标志和标题，并核对文号、日期。例如：某政发〔2026〕1号、2026年9月12日。";
                return;
            }
            if ((要素.印发机关.Length > 0) != (要素.印发日期.Length > 0) || (要素.印发日期.Length > 0 && !公文要素服务.日期有效(要素.印发日期)))
            {
                error.Text = "印发机关和印发日期需同时填写，日期必须有效。"; return;
            }
            要素.预留盖章 = stamp.IsChecked == true;
            要素.已确认 = true;
            DialogResult = true;
        };
    }
}
