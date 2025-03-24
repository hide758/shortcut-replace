"""ショートカット変換スクリプト
Windowsのショートカットのリンク先を変更するスクリプトです。
指定したフォルダ以下のショートカットファイル(*.lnk)を探し、リンク先のパスに指定した文字列が含まれている場合、指定した文字列を別の文字列に置換します。

Useage:
    python main.py --original-path-string "置換前の文字列" --replace-path-string "置換後の文字列" "ショートカットを探すフォルダ" [--dry-run]
"""

from pathlib import Path
import argparse
import re
import csv
import win32com.client 

default_original_path = "\\\\mei660\\Dsi\\SV_NC"
default_replace_path = "\\\\mei660\\NC\\NC_NCDrive\\"

if __name__ == '__main__' :
    try:
        # parse arguments
        parser = argparse.ArgumentParser(description='Windowsショートカットのリンク先を変更するスクリプト')

        parser.add_argument('--original-path-string', default=default_original_path, help='ショートカットパスの文字列(置換前)')
        parser.add_argument('--replace-path-string', default=default_replace_path, help='ショートカットパスの文字列(置換後)')
        parser.add_argument('target-directory', default=".", help='ショートカットを探すフォルダ')
        parser.add_argument('--dry-run', help='実際には変更を行わない', action='store_true')

        args = parser.parse_args()

        # get arguments
        src_str = args.original_path_string
        dst_str = args.replace_path_string
        targetdir = Path(vars(args)["target-directory"])

        shell = win32com.client.Dispatch("WScript.Shell")

        print("search shortcut files...")
        report = []
        LinkList = list(targetdir.glob("**/*.lnk"))

        # confirm convert
        print(f"{len(LinkList)} shortcut files found.")
        print("ショートカットを変換しますか？(y/n)")
        if input() != "y":
            print("変換を中止しました。")
            exit(0)
            
        for cnt, shortcutpath in enumerate(LinkList, 1):
            print(f"\n[{cnt:3d} / {len(LinkList):3d}] inspect {shortcutpath}")
            try:
                # create shortcut object
                shortcut = shell.CreateShortCut(str(shortcutpath))

                TargetPath = shortcut.TargetPath
                WorkPath = shortcut.WorkingDirectory

                rep = {
                    "Convert" : False,
                    "LinkPath" : str(shortcutpath),
                    "BeforeTargetPath" : TargetPath,
                    "BeforeWorkPath" : WorkPath,
                    "AfterTargetPath" : None,
                    "AfterWorkPath" : None,
                }

                
                # target path
                if re.search(re.escape(src_str), TargetPath, flags=re.IGNORECASE):
                    # convert path
                    renew_target = str(Path(re.sub(re.escape(src_str), lambda m : dst_str, TargetPath, flags=re.IGNORECASE)))

                    # replace link path
                    shortcut.TargetPath = renew_target
                    rep["AfterTargetPath"] = renew_target

                # work path
                if re.search(re.escape(src_str), WorkPath, flags=re.IGNORECASE):
                    # convert path
                    renew_work = str(Path(re.sub(re.escape(src_str), lambda m : dst_str, WorkPath, flags=re.IGNORECASE)))

                    # replace work path
                    shortcut.WorkingDirectory = renew_work
                    rep["AfterWorkPath"] = renew_work

                # save shortcut
                if rep["AfterTargetPath"] != None or rep["AfterWorkPath"] != None:
                    print("  ==> convert")
                    rep["Convert"] = True

                    # save shortcut
                    if vars(args)["dry_run"] == False:
                        shortcut.Save()

                # add report
                report.append(rep)

            except Exception as e:
                print(f"  {e.args[1]}\n  {hex(e.args[0] & 0xffffffff)}")
                continue

        
        # export report.csv
        header = ['No.', '変換', 'ショートカット', '変換前 リンク先', '変換前 作業フォルダー', '変換後 リンク先', '変換後 作業フォルダー']
        with open('report.csv', 'w', newline='', encoding="utf_8_sig") as f:
            writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header)
            for no, row in enumerate(report, 1):
                writer.writerow([
                    no,
                    "" if row["Convert"] == False else "✓",
                    row["LinkPath"],
                    row["BeforeTargetPath"],
                    row["BeforeWorkPath"],
                    row["AfterTargetPath"] if row["AfterTargetPath"] != None else "(変換なし)",
                    row["AfterWorkPath"] if row["AfterWorkPath"] != None else "(変換なし)",
                    ])

    except Exception as e:
        print(e)
