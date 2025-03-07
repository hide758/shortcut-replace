# shortcut-replace
A tool for batch replacing target paths in Windows shortcuts

Windowsショートカットのリンク先を一括置換するツール


## Overview
This tool batch replaces the target paths of Windows shortcut (.lnk) files within a specified directory.

It can be used, for example, to update shortcut targets when a file server is changed.

本ツールは、指定したディレクトリ内のWindowsショートカット（.lnkファイル）のリンク先パスを一括で置換します。
ファイルサーバー変更などで、リンク先修正する場合に使用可能です。


## Setup

Python3.x環境をインストールしてください。次に、以下のコマンドで必要なパッケージをインストールします。  

Please install a Python 3.x environment. Then, install the required packages using the command below.

```bash
pip install -r requirements.txt
```

## Usage Examples
```bash
python main.py "\\srv01\folder1\" "\\srv02\folder2\" "c:\myshortcuts"
```

## Command-Line

```bash
python main.py [String to be replaced] [Replacement string] [Target directory] [--dry-run]
```

- String to be replaced (置換前の文字列) : The string to search for. (置換対象となる文字列)

- Replacement string (置換後の文字列) : The string to replace with. (置換後に使用する文字列)

- Target directory (対象ディレクトリ) : The directory where shortcuts (.lnk files) are located. (ショートカットが含まれるディレクトリのパス)

- --dry-run : Optional flag to perform a dry run without making actual changes.
(実際の変更を行わず、処理内容を確認するためのオプション) 
