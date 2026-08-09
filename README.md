# GB/T 9704—2012 排版工具

这是一个 Windows 桌面程序，用 Open XML 直接处理 `.docx` 文件，按 GB/T 9704—2012 的常用公文版式统一页面、正文、标题、表格、页眉页脚和页码。

程序支持处理单个文件或整个文件夹。原文件不会被覆盖，结果默认保存到原目录，也可以另选输出目录。

## 能做什么

- 自动识别普通公文、信函、命令和纪要，也可以在界面中手动指定文种。
- 统一 A4 页面尺寸、页边距、版心栅格、正文段落和多级标题格式。
- 处理表格文字、数字、边框和对齐方式。
- 处理页眉、页脚、页码字段、附件提示和版记区域。
- 批量处理文件夹，可选择是否包含子文件夹。
- 每个输出文件旁生成一份 `.audit.json` 审计记录；批量处理另有汇总清单。
- 对图片、超链接、书签等复杂结构采取保守处理，并在日志或审计记录中提示。

## 使用限制

- 只处理 `.docx`，不处理旧版 `.doc`。
- 只支持 Windows。
- 文种识别负责选择现有通用排版流程；信函、命令和纪要尚没有各自独立的专属重排模块。
- 复杂模板、宏、嵌入对象和非标准域代码请先用副本试跑并人工复核。

## 直接运行

需要安装 [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)。

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
2. 按需选择输出目录、子文件夹和文种模式。
3. 点击“开始处理”。再次点击可以请求取消尚未完成的批次。
4. 查看界面日志、输出文件和审计 JSON。

输出文件名以 `_GB9704-2012排版.docx` 结尾。程序会跳过已经带有该后缀的文件。

## 构建和测试

```powershell
dotnet restore .\GB-T9704-2012排版工具.csproj
dotnet build .\GB-T9704-2012排版工具.csproj -c Release
dotnet test .\GB-T9704-2012排版工具.Tests\GB-T9704-2012排版工具.Tests.csproj -c Release
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
