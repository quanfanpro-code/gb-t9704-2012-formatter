using System.Windows.Media;

namespace GBT9704_2012排版工具.Core;

public sealed class GovFontCheckService
{
    private static readonly string[] 必需字体 =
    [
        "方正小标宋简体",
        "仿宋_GB2312",
        "楷体_GB2312",
        "黑体",
        "宋体"
    ];

    public IReadOnlyList<string> 获取缺失字体()
    {
        var installed = Fonts.SystemFontFamilies
            .SelectMany(x => x.FamilyNames.Values.Append(x.Source))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return 必需字体.Where(x => !installed.Contains(x)).ToList();
    }
}

