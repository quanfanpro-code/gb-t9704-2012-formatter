using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovHeaderFooterService
{
    // 页面常量统一引用 GovPageConstants，避免与校验服务双份硬编码

    public void 格式化(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("文档缺少主部件。");
        var body = mainPart.Document.Body ?? throw new InvalidOperationException("文档缺少正文。");

        // 清除旧版 Word 可能塞在 SectionProperties 中的 EvenAndOddHeaders（该元素只允许出现在 Settings 中）
        清除非法EvenAndOddHeaders(document);

        // 全局启用奇偶页不同：FooterReference type=Even 只有在 settings.xml 设置 w:evenAndOddHeaders 后才会被 Word 采用，
        // 否则双数页也会套用 Default 页脚（右对齐页码），违反 GB/T 9704 双页页码居左空一字的要求
        确保奇偶页设置(mainPart);

        foreach (var section in 构建分节(body))
        {
            配置页面(section.SectionProperties);
            复制偶数页页眉(section.SectionProperties, mainPart);

            if (是版记节(section))
                continue;

            重建页脚(section.SectionProperties, mainPart);
        }
    }

    private static void 配置页面(SectionProperties sectionProperties)
    {
        // Schema 顺序：HeaderReference → FooterReference → footnotePr/endnotePr/type → PageSize → PageMargin → ... → DocGrid → PrinterSettings
        var pageSize = sectionProperties.GetFirstChild<PageSize>();
        if (pageSize == null)
        {
            pageSize = new PageSize();
            // 先找最后一个 HeaderReference/FooterReference，插入到其后
            var afterHeaders = sectionProperties.ChildElements
                .LastOrDefault(x => x is HeaderReference or FooterReference);
            if (afterHeaders != null)
            {
                sectionProperties.InsertAfter(pageSize, afterHeaders);
            }
            else
            {
                // 无页眉页脚引用时不能简单 PrependChild：footnotePr/endnotePr/type 按 schema 应排在 PageSize 之前
                var anchor = sectionProperties.ChildElements.FirstOrDefault(x =>
                    x.LocalName is not "footnotePr" and not "endnotePr" and not "type");
                if (anchor != null)
                    sectionProperties.InsertBefore(pageSize, anchor);
                else
                    sectionProperties.AppendChild(pageSize);
            }
        }

        var isLandscape = pageSize.Orient?.Value == PageOrientationValues.Landscape ||
                          (pageSize.Width?.Value > pageSize.Height?.Value);

        if (isLandscape)
        {
            pageSize.Width = (UInt32Value)(uint)GovPageConstants.A4高度;
            pageSize.Height = (UInt32Value)(uint)GovPageConstants.A4宽度;
            pageSize.Orient = PageOrientationValues.Landscape;
        }
        else
        {
            pageSize.Width = (UInt32Value)(uint)GovPageConstants.A4宽度;
            pageSize.Height = (UInt32Value)(uint)GovPageConstants.A4高度;
            pageSize.Orient = PageOrientationValues.Portrait;
        }

        // PageMargin 必须在 PageSize 之后
        var pageMargin = sectionProperties.GetFirstChild<PageMargin>();
        if (pageMargin == null)
        {
            pageMargin = new PageMargin();
            sectionProperties.InsertAfter(pageMargin, pageSize);
        }

        pageMargin.Top = (Int32Value)(int)GovPageConstants.上边距;
        pageMargin.Bottom = (Int32Value)(int)GovPageConstants.下边距;
        pageMargin.Left = GovPageConstants.左边距;
        pageMargin.Right = GovPageConstants.右边距;
        pageMargin.Header = (UInt32Value)(uint)GovPageConstants.页眉距离;
        pageMargin.Footer = (UInt32Value)(uint)GovPageConstants.页脚距离;
        pageMargin.Gutter = 0U;

        // DocGrid 必须在 PrinterSettingsReference 之前
        var docGrid = sectionProperties.GetFirstChild<DocGrid>();
        if (docGrid == null)
        {
            docGrid = new DocGrid();
            var printerSettings = sectionProperties.ChildElements
                .FirstOrDefault(x => x.LocalName == "printerSettings" || x.LocalName == "printerSettingsReference");
            if (printerSettings != null)
                sectionProperties.InsertBefore(docGrid, printerSettings);
            else
                sectionProperties.AppendChild(docGrid);
        }

        docGrid.Type = DocGridValues.Lines;
        docGrid.LinePitch = GovPageConstants.正文行距;
    }

    /// <summary>
    /// 确保 settings.xml 中存在 w:evenAndOddHeaders，奇偶页页脚引用才会被 Word 采用。
    /// 没有 DocumentSettingsPart 时新建；SDK 3.x 的 Settings 没有该元素的强类型属性，
    /// 用 AddChild 由 SDK 按 schema 粒子顺序插入，不裸 AppendChild。
    /// </summary>
    private static void 确保奇偶页设置(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        var settings = settingsPart.Settings;
        if (settings.GetFirstChild<EvenAndOddHeaders>() == null)
            settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
    }

    private static void 复制偶数页页眉(SectionProperties sectionProperties, MainDocumentPart mainPart)
    {
        var headerRefs = sectionProperties.Elements<HeaderReference>().ToList();
        if (headerRefs.Any(x => x.Type?.Value == HeaderFooterValues.Even))
            return;

        var oddHeader = headerRefs.FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.Default)
                        ?? headerRefs.FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.First)
                        ?? headerRefs.FirstOrDefault();
        if (oddHeader?.Id == null)
            return;

        插入节引用(sectionProperties, new HeaderReference
        {
            Type = HeaderFooterValues.Even,
            Id = oddHeader.Id
        });
    }

    private static void 重建页脚(SectionProperties sectionProperties, MainDocumentPart mainPart)
    {
        // 保留首页页脚（type=First）引用，只重建 Default/Even 页脚
        var 待清理关系 = new List<string>();
        foreach (var footerRef in sectionProperties.Elements<FooterReference>().ToList())
        {
            if (footerRef.Type?.Value == HeaderFooterValues.First)
                continue;

            if (!string.IsNullOrWhiteSpace(footerRef.Id?.Value))
                待清理关系.Add(footerRef.Id!.Value!);
            footerRef.Remove();
        }

        // 仅当旧页脚部件不再被文档中任何节的 FooterReference 引用时才删除，避免误删其他节共享的部件
        foreach (var relationId in 待清理关系.Distinct())
        {
            try
            {
                var 仍被引用 = mainPart.Document.Descendants<FooterReference>()
                    .Any(x => x.Id?.Value == relationId);
                if (仍被引用)
                    continue;

                var orphanPart = mainPart.FooterParts
                    .FirstOrDefault(fp => mainPart.GetIdOfPart(fp) == relationId);
                if (orphanPart != null)
                    mainPart.DeletePart(orphanPart);
            }
            catch (InvalidOperationException ex)
            {
                // 删除孤立页脚部件失败时不做阻断，避免排版流程因清理逻辑中断
                System.Diagnostics.Debug.WriteLine($"删除孤立页脚部件 {relationId} 失败：{ex.Message}");
            }
        }

        var oddFooterPart = mainPart.AddNewPart<FooterPart>();
        oddFooterPart.Footer = new Footer(创建页脚段落(JustificationValues.Right, rightChars: 1));

        var evenFooterPart = mainPart.AddNewPart<FooterPart>();
        evenFooterPart.Footer = new Footer(创建页脚段落(JustificationValues.Left, leftChars: 1));

        插入节引用(sectionProperties, new FooterReference
        {
            Type = HeaderFooterValues.Default,
            Id = mainPart.GetIdOfPart(oddFooterPart)
        });

        插入节引用(sectionProperties, new FooterReference
        {
            Type = HeaderFooterValues.Even,
            Id = mainPart.GetIdOfPart(evenFooterPart)
        });
    }

    private static Paragraph 创建页脚段落(JustificationValues alignment, int leftChars = 0, int rightChars = 0)
    {
        var paragraph = new Paragraph();
        GovOpenXmlHelper.设置段落行距(paragraph, 420);
        GovOpenXmlHelper.设置段落缩进(paragraph, leftChars: leftChars, rightChars: rightChars);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = alignment };
        GovOpenXmlHelper.添加页码域(paragraph, "Times New Roman", "28");

        return paragraph;
    }

    private static bool 是版记节(GovSectionContext section)
    {
        var text = string.Join(Environment.NewLine, section.Texts).Trim();
        if (string.IsNullOrWhiteSpace(text))
            return true;

        if (section.Texts.Count > 6)
            return false;

        return text.Contains("抄送", StringComparison.Ordinal) ||
               text.Contains("印发", StringComparison.Ordinal) ||
               text.Contains("主送", StringComparison.Ordinal);
    }

    private static List<GovSectionContext> 构建分节(Body body)
    {
        var sections = new List<GovSectionContext>();
        var texts = new List<string>();
        var currentSectPr = body.GetFirstChild<SectionProperties>();

        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph paragraph)
            {
                texts.Add(GovOpenXmlHelper.段落文本(paragraph).Trim());
                var paraSectPr = paragraph.ParagraphProperties?.SectionProperties;
                if (paraSectPr != null)
                {
                    sections.Add(new GovSectionContext(paraSectPr, texts.Where(x => !string.IsNullOrWhiteSpace(x)).ToList()));
                    texts = [];
                }
            }
            else if (element is Table table)
            {
                texts.Add(table.InnerText?.Trim() ?? string.Empty);
            }
            else if (element is SectionProperties bodySectPr)
            {
                currentSectPr = bodySectPr;
                sections.Add(new GovSectionContext(bodySectPr, texts.Where(x => !string.IsNullOrWhiteSpace(x)).ToList()));
                texts = [];
            }
        }

        if (sections.Count == 0 && currentSectPr != null)
        {
            sections.Add(new GovSectionContext(currentSectPr, texts.Where(x => !string.IsNullOrWhiteSpace(x)).ToList()));
        }

        return sections.DistinctBy(x => x.SectionProperties).ToList();
    }

    private static void 插入节引用(SectionProperties sectionProperties, OpenXmlElement reference)
    {
        var anchor = sectionProperties.ChildElements.FirstOrDefault(x =>
            x is not HeaderReference &&
            x is not FooterReference);

        if (anchor == null)
            sectionProperties.AppendChild(reference);
        else
            sectionProperties.InsertBefore(reference, anchor);
    }

    /// <summary>
    /// 遍历文档所有部件（正文、页眉、页脚），删除 Settings 之外的所有 EvenAndOddHeaders。
    /// 旧版 Word 兼容模式下可能把它塞进 SectionProperties，严格 Schema 不允许。
    /// </summary>
    private static void 清除非法EvenAndOddHeaders(WordprocessingDocument document)
    {
        // 主文档
        var mainDoc = document.MainDocumentPart?.Document;
        if (mainDoc != null)
        {
            foreach (var e in mainDoc.Descendants<EvenAndOddHeaders>().ToList())
                e.Remove();
        }

        // 页眉部件
        if (document.MainDocumentPart?.HeaderParts != null)
        {
            foreach (var headerPart in document.MainDocumentPart.HeaderParts)
            {
                foreach (var e in headerPart.Header.Descendants<EvenAndOddHeaders>().ToList())
                    e.Remove();
            }
        }

        // 页脚部件
        if (document.MainDocumentPart?.FooterParts != null)
        {
            foreach (var footerPart in document.MainDocumentPart.FooterParts)
            {
                foreach (var e in footerPart.Footer.Descendants<EvenAndOddHeaders>().ToList())
                    e.Remove();
            }
        }
    }
}

public sealed record GovSectionContext(SectionProperties SectionProperties, List<string> Texts);

