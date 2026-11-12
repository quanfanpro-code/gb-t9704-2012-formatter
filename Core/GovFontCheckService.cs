using System.Windows.Media;

namespace GBT9704_2012排版工具.Core;

public sealed class GovFontCheckService
{
    // 国标标题字体为方正小标宋_GBK，未安装时退回方正小标宋简体，保证老机器仍能排版
    private const string 标题字体主选 = "方正小标宋_GBK";
    private const string 标题字体备选 = "方正小标宋简体";

    private static readonly string[] 其它必需字体 =
    [
        "仿宋_GB2312",
        "楷体_GB2312",
        "黑体",
        "宋体",
        "Times New Roman"
    ];

    public static string 选定标题字体() =>
        已安装字体().Contains(标题字体主选) ? 标题字体主选 : 标题字体备选;

    public IReadOnlyList<string> 获取缺失字体()
    {
        var installed = 已安装字体();

        var 缺失 = 其它必需字体.Where(x => !installed.Contains(x)).ToList();
        if (!installed.Contains(标题字体主选) && !installed.Contains(标题字体备选))
            缺失.Insert(0, $"{标题字体主选}（或{标题字体备选}）");

        return 缺失;
    }

    private static HashSet<string> 已安装字体() =>
        Fonts.SystemFontFamilies
            .SelectMany(x => x.FamilyNames.Values.Append(x.Source))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
}

