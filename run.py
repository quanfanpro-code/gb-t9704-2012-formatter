from pathlib import Path
import subprocess


项目文件 = Path(__file__).with_name("GB-T9704-2012排版工具.csproj")
subprocess.Popen(["dotnet", "run", "--project", str(项目文件)], cwd=项目文件.parent)
