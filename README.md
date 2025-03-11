# LiteratureInterpretation
 
### 1. 安装依赖库

确保您已经安装了所有必要的依赖库。您可以通过以下命令安装：
```
pip install pdfplumber pandas requests openpyxl tkinter
```

###  2. 下载 PyInstaller

PyInstaller 是用于打包 Python 应用程序的工具。您可以使用以下命令安装：
```
pip install pyinstaller
```

### 3. 打包程序
使用以下命令打包您的程序：
```
pyinstaller --onefile --windowed your_script.py
```
--onefile：将所有文件打包成一个可执行文件。
--windowed：隐藏控制台窗口，仅显示 GUI。


### 4. 运行可执行文件

打包完成后，您可以在 dist 文件夹中找到生成的可执行文件。您可以将其分发给其他用户，只需双击即可运行。
