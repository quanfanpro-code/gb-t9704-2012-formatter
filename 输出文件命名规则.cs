using System.IO;

namespace GBT9704_2012排版工具;

public static class 输出文件命名规则
{
    private const string 输出后缀 = "_GB9704-2012排版";

    public static string 生成输出路径(string inputPath)
    {
        return 生成输出路径(inputPath, null);
    }

    public static string 生成输出路径(string inputPath, string? outputDirectory)
    {
        var dir = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.GetDirectoryName(inputPath)!
            : outputDirectory;
        var name = Path.GetFileNameWithoutExtension(inputPath);
        var path = Path.Combine(dir, $"{name}{输出后缀}.docx");
        if (!File.Exists(path)) return path;
        for (var i = 2; i < 100; i++)
        {
            var candidate = Path.Combine(dir, $"{name}{输出后缀}({i}).docx");
            if (!File.Exists(candidate)) return candidate;
        }
        return Path.Combine(dir, $"{name}{输出后缀}_{Guid.NewGuid():N}.docx");
    }

    public static bool 是已排版文件(string path)
    {
        // 以"主文件名以后缀结尾"判断，避免原文件名碰巧包含后缀文本被误判跳过
        return Path.GetFileNameWithoutExtension(path)
            .EndsWith(输出后缀, StringComparison.OrdinalIgnoreCase);
    }
}

