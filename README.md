# Convert-system-information-to-CSV

This is Python version 3

![License](https://img.shields.io/pypi/pyversions/3)


Using Python
====
Installable Python kits, and information about using Python, are available at [python.org](https://www.python.org/)。


📦Install
pip automatic download & update
-------
這段程式中已經使用了幾個第三方庫，但還有一個 tqdm 库，它是用來在循環中顯示進度條的。
如果你的系統中沒有安裝它，你需要使用 pip 來安裝它。
```
pip install tqdm
```
注意
msvcrt 是一个Windows特定的库，用于文件锁定，如果你的程序在其他平台上运行，可能会导致问题。

另外，如果你想在 Windows 上使用文件锁定，msvcrt 库是一个选择，但在其他平台上可能需要使用不同的方法来实现文件锁定。

📦Pack
-------
```
pip install pyinstaller
```
```
pyinstaller –F + Project name.py
```
This will generate a folder named dist in the current directory containing the packaged executable file.  

If you need the progress bar version, you can go to
