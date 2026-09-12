using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GBT9704_2012排版工具.Contracts;

namespace GBT9704_2012排版工具.Core;

public static class 公文要素服务
{
    internal static string 文本(Paragraph p) => GovOpenXmlHelper.提取可见文本(p).Trim();
    internal static bool 是文号候选(string s) => Regex.IsMatch(s, @"^.{1,30}[〔\[（(]?\d{4}[〕\]）)].*号$");
    internal static bool 是日期候选(string s) => Regex.IsMatch(s, @"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$");
    public static bool 日期有效(string s) => DateOnly.TryParseExact(s, "yyyy年M月d日", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
        && d.ToString("yyyy年M月d日", CultureInfo.InvariantCulture) == s;
    public static bool 文号有效(string s) => Regex.IsMatch(s, @"^[\p{IsCJKUnifiedIdeographs}A-Za-z]+〔[1-9]\d{3}〕[1-9]\d*号$");

    public static 公文要素 读取(MainDocumentPart main)
    {
        var ps = (main.Document?.Body ?? throw new InvalidOperationException("文档缺少正文。")).Elements<Paragraph>().ToArray();
        var ts = ps.Select(文本).ToArray();
        var structure = new GovStructureAnalyzer().分析(main);
        var number = Array.FindIndex(ts, 是文号候选);
        var date = Array.FindLastIndex(ts, 是日期候选);
        var title = structure.标题段索引.Order().FirstOrDefault(-1);
        var agency = number > 0 ? ts.Take(number).LastOrDefault(s => s.Length > 0) ?? "" : "";
        return new 公文要素
        {
            机关标志 = agency,
            标题 = title >= 0 ? ts[title] : "",
            文号 = number >= 0 ? ts[number] : "",
            主送机关 = structure.主送机关段索引.Order().Select(i => ts[i]).FirstOrDefault() ?? "",
            成文日期 = date >= 0 ? ts[date] : "",
            署名 = date > 0 && ts[date - 1].Length is > 0 and <= 40 && !Regex.IsMatch(ts[date - 1], "[。；：:]") ? ts[date - 1] : "",
            抄送 = ts.FirstOrDefault(s => s.StartsWith("抄送")) ?? "",
            印发机关 = ts.Where(s => s.EndsWith("印发")).Select(s => Regex.Split(s, @"\d{4}年")[0].Trim()).FirstOrDefault() ?? "",
            印发日期 = ts.Where(s => s.EndsWith("印发")).Select(s => Regex.Match(s, @"\d{4}年\d{1,2}月\d{1,2}日").Value).FirstOrDefault() ?? ""
        };
    }

    public static List<string> 检查(MainDocumentPart main, RequestContract request)
    {
        var texts = (main.Document?.Body ?? throw new InvalidOperationException("文档缺少正文。")).Elements<Paragraph>().Select(文本).Where(s => s.Length > 0).ToList();
        var errors = new List<string>();
        var formal = request.模式 != 排版模式.普通材料;
        var f = request.要素;
        var numbers = texts.Take(20).Where(是文号候选).ToArray();
        var dates = texts.Where(是日期候选).ToArray();
        if (formal && f is not null)
        {
            numbers = [f.文号];
            dates = new[] { f.成文日期, f.印发日期 }.Where(s => s.Length > 0).ToArray();
        }
        if (numbers.Any(s => !文号有效(s))) errors.Add("文号格式错误：应为机关代字〔年份〕序号号，不加前导零或‘第’字。");
        if (texts.Take(20).Count(是文号候选) > 1 && f?.已确认 != true) errors.Add("原稿存在多个文号，请逐文件核对并明确使用的文号。");
        if (dates.Any(s => !日期有效(s))) errors.Add("日期格式或日历日期无效，请使用无前导零的年月日。");
        if (formal)
        {
            if (f is null || string.IsNullOrWhiteSpace(f.标题) || string.IsNullOrWhiteSpace(f.机关标志) || string.IsNullOrWhiteSpace(f.文号) || string.IsNullOrWhiteSpace(f.成文日期))
                errors.Add("正式公文缺少机关标志、标题、文号或成文日期，请补充。");
            if (request.模式 == 排版模式.套打 && (!double.IsFinite(request.套打红线距纸顶毫米) || request.套打红线距纸顶毫米 is < 70 or > 150))
                errors.Add("套打红线高度须在距纸顶 70 至 150 毫米之间，超出后首页正文空间不足。");
            if (f is not null && (f.印发机关.Length > 0) != (f.印发日期.Length > 0))
                errors.Add("印发机关和印发日期需同时填写。");
        }
        request.检查记录.Add(new("ELEMENTS", "文号和日期", errors.Count == 0 ? "通过" : "需复核", errors.Count == 0 ? "已核对存在的要素；未提供的可选项目不适用。" : string.Join("\n", errors)));
        var attachmentErrors = 检查附件(texts);
        errors.AddRange(attachmentErrors);
        request.检查记录.Add(new("ATTACHMENTS", "附件一致性", attachmentErrors.Count > 0 ? "需复核" : texts.Any(s => s.StartsWith("附件")) ? "通过" : "不适用", string.Join("\n", attachmentErrors)));
        return errors;
    }

    private static List<string> 检查附件(List<string> texts)
    {
        var declarations = new List<(int 编号, string 标题)>();
        var labels = new List<(int 编号, string 标题)>();
        var continuation = false;
        for (var i = 0; i < texts.Count; i++)
        {
            var t = texts[i];
            var d = Regex.Match(t, @"^附件[：:]\s*(?:(\d+)[.．、]\s*)?(.+)$");
            var c = Regex.Match(t, @"^(\d+)[.．、]\s*(.+)$");
            if (d.Success)
            {
                declarations.Add((d.Groups[1].Success ? int.Parse(d.Groups[1].Value) : 1, d.Groups[2].Value.Trim()));
                continuation = true;
            }
            else if (continuation && c.Success)
                declarations.Add((int.Parse(c.Groups[1].Value), c.Groups[2].Value.Trim()));
            else continuation = false;
            var label = Regex.Match(t, @"^附件\s*(\d*)$");
            if (label.Success)
                labels.Add((label.Groups[1].Value.Length > 0 ? int.Parse(label.Groups[1].Value) : 1, i + 1 < texts.Count ? texts[i + 1] : ""));
        }
        if (declarations.Count == 0 && labels.Count == 0) return [];
        if (declarations.Count != labels.Count) return ["附件说明与所附附件数量不一致，或正文未包含全部附件，请人工核对。"];
        for (var i = 0; i < declarations.Count; i++)
            if (declarations[i].编号 != i + 1 || labels[i].编号 != i + 1 || declarations[i].标题 != labels[i].标题)
                return ["附件说明与附件标签的顺序或标题不一致，请核对。"];
        return [];
    }
}
