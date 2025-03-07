from pathlib import Path
import argparse
import win32com.client 

default_original_path = "\\mei660/Dsi"
default_replace_path = "\\mei660/NC"

if __name__ == '__main__' :
    # parse arguments
    parser = argparse.ArgumentParser(description='Windowsショートカットのリンク先を変更するスクリプト')

    parser.add_argument('original-path-string', default=default_original_path, help='ショートカットパスの文字列(置換前)')
    parser.add_argument('replace-path-string', default=default_replace_path, help='ショートカットパスの文字列(置換後)')
    parser.add_argument('target-directory', default=".", help='ショートカットが存在するフォルダ')
    parser.add_argument('--dry-run', help='実際には変更を行わない', action='store_true')

    args = parser.parse_args()

    # get arguments
    src_str = Path(vars(args)["original-path-string"])
    dst_str = Path(vars(args)["replace-path-string"])
    targetdir = Path(vars(args)["target-directory"])

    shell = win32com.client.Dispatch("WScript.Shell")

    for shortcut in targetdir.glob("**/*.lnk"):
        # create shortcut object
        shortcut = shell.CreateShortCut(str(shortcut))

        # target path
        if str(src_str) in str(Path(shortcut.Targetpath)):
            print(f"\r\n*** {str(shortcut)}")
            print(f"before : {str(shortcut.Targetpath)} \r\n after : {str(shortcut.Targetpath).replace(str(src_str), str(dst_str))}")

            if vars(args)["dry_run"] == False:
                # replace link path
                shortcut.Targetpath=shortcut.Targetpath.replace(str(src_str), str(dst_str))
                shortcut.save()
