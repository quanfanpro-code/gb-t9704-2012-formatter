# GB/T 9704—2012 排版工具

这是一个 Windows 桌面程序，用 Open XML 直接处理 `.docx` 文件，支持普通材料、正式公文电子红头和预印红头纸套打。

程序支持处理单个文件或整个文件夹。原文件不会被覆盖，结果默认保存到原目录，也可以另选输出目录。

## 能做什么

- 默认处理普通材料；正式公文逐份核对机关标志、标题、文号、主送、署名日期及版记。
- 电子红头生成机关标志和红线；套打不重复绘制红头，可填写纸顶至预印红线的距离。
- 正式公文支持附件另页、盖章空间及末页版记；长标题和机关名称可按词意手动分行。
- 统一 A4 页面尺寸、页边距、版心栅格、正文段落和多级标题格式。
- 保留表格文字、数字、行列、图片和对象，不擅自解释业务数据。
- 保留原页眉；各节旧页脚统一替换为国标页码。
- 批量处理文件夹，可选择是否包含子文件夹。
- 每个输出文件旁生成 `.audit.json` 检查记录；批量处理另有汇总清单。
- 保留公式、批注、修订、脚注、尾注、文本框和嵌入对象；相关段落跳过自动改写并提示人工确认。
- 在最终输出通过 OpenXML 结构校验后才发布结果，不覆盖同名文件。

## 使用限制

- 只处理 `.docx`，不处理旧版 `.doc`。
- 只支持 Windows。
- 本版本不处理上行文，也不提供信函、命令、纪要专属格式。
- 机关标志、发文字号或标题无法可靠识别时，程序采用保守处理，不能代替人工定稿。
- Open XML 结构校验只能证明文件结构有效，不能证明 Word、WPS 中的每一页都达到像素级一致；正式发文前仍应查看版面。
- `.docm` 宏文档、加密文档和损坏文档不在支持范围内。

## 直接运行

需要安装 [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)。

在项目目录运行：

```powershell
python .\run.py
```

也可以直接运行：

```powershell
dotnet run --project .\GB-T9704-2012排版工具.csproj
```

打开程序后：

1. 选择一个 Word 文件或文件夹。
2. 选择普通材料、正式公文电子红头或预印红头纸套打，再选择输出目录和是否包含子文件夹。
3. 点击“开始处理”。正式公文需逐份结合原文核对要素；套打先填写实际纸样的红线高度。再次点击可以请求取消尚未完成的批次。
4. 查看界面日志、输出文件和审计 JSON。

输出文件名以 `_GB9704-2012排版.docx` 结尾。名称冲突时自动增加序号，程序不会覆盖已有文件。处理失败时保留带“失败”标识的副本供排查，原文件不受影响。

检查结果分为三类：

- `通过`：最终文件结构有效，已执行的检查没有发现问题。
- `人工复核`：文件已成功生成，但复杂内容或不确定要素需要人工查看。
- `阻断`：最终文件结构无效、无法打开或处理失败，没有发布正常结果。

## 构建和测试

2026-09-12：已使用 Windows Microsoft Word 16.0.20326 和真实字体核验 18 份样本、30 页输出；小标宋为 4.00 版，嵌入标记为 8（允许可编辑嵌入）。覆盖普通材料、电子红头、两种套打高度、四级标题、长标题、附件、版记及复杂内容保留。该结果不代表 WPS 或实体打印机套打已经实测。程序日常生成的检查记录会明确标注 Word 实测未执行，不会冒充逐份人工验收。

```powershell
dotnet restore .\GB-T9704-2012排版工具.csproj
dotnet build .\GB-T9704-2012排版工具.csproj -c Release
dotnet test .\GB-T9704-2012排版工具.Tests\GB-T9704-2012排版工具.Tests.csproj -c Release
dotnet format .\GB-T9704-2012排版工具.csproj --verify-no-changes --no-restore
dotnet format .\GB-T9704-2012排版工具.Tests\GB-T9704-2012排版工具.Tests.csproj --verify-no-changes --no-restore
dotnet list .\GB-T9704-2012排版工具.csproj package --vulnerable --include-transitive
```

发布为 Windows x64 单文件程序：

```powershell
dotnet publish .\GB-T9704-2012排版工具.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o .\publish
```

发布目录是生成物，不应提交到源码仓库。

## 主要目录

```text
Contracts/                         请求、响应和审计数据结构
Core/                              文档识别、排版、校验和审计
GB-T9704-2012排版工具.Tests/       xUnit 回归测试
MainWindow.xaml                    桌面界面
run.py                             本地启动脚本
```

## 数据和隐私

文档只在本机处理，程序没有上传功能。公开提交或反馈问题时，请勿附带真实公文、审计 JSON、客户名称、账号、联系方式或其他敏感信息。仓库默认忽略 Word 文档和本地生成物。

## 许可证

本项目采用 [GNU Affero General Public License v3.0](LICENSE)，许可证标识为 `AGPL-3.0-only`。

Open XML SDK 和 WPF-UI 使用 MIT 许可证；测试依赖使用 MIT 或 Apache-2.0 许可证。
