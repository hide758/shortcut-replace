from pathlib import Path
import argparse
import re
import csv
import win32com.client 

default_original_path = "\\mei660/Dsi"
default_replace_path = "\\mei660/NC"

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
        for cnt, shortcut in enumerate(LinkList, 1):
            # create shortcut object
            shortcut = shell.CreateShortCut(str(shortcut))
            TargetPath = shortcut.TargetPath
            WorkPath = shortcut.WorkingDirectory

            rep = {
                "LinkPath" : str(shortcut),
                "BeforeTargetPath" : TargetPath,
                "BeforeWorkPath" : WorkPath,
                "AfterTargetPath" : None,
                "AfterWorkPath" : None,
            }

            print(f"[{cnt:3d} / {len(LinkList):3d}] inspect {rep['LinkPath']}")
            
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

            # add report
            if rep["AfterTargetPath"] != None or rep["AfterWorkPath"] != None:
                print("  ==> convert")
                report.append(rep)

                # save shortcut
                if vars(args)["dry_run"] == False:
                    shortcut.Save()

        
        # export report.csv
        header = ['No.', 'ショートカット', '変換前 リンク先', '変換前 作業フォルダー', '変換後 リンク先', '変換後 作業フォルダー']
        with open('report.csv', 'w', newline='') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header)
            for no, row in enumerate(report, 1):
                writer.writerow([
                    no,
                    row["LinkPath"],
                    row["BeforeTargetPath"],
                    row["BeforeWorkPath"],
                    row["AfterTargetPath"] if row["AfterTargetPath"] != None else "(変換なし)",
                    row["AfterWorkPath"] if row["AfterWorkPath"] != None else "(変換なし)",
                    ])

    except Exception as e:
        print(e)
