# Copyright (c) 2022 sangoopan
# Released under the MIT license
# https://licenses.opensource.jp/MIT/MIT.html

import csv
import os
import sys
from tkinter import StringVar, Tk, W, filedialog, ttk


class DialogTitle:
    MAIN_DIALOG = "SRPG Studio用 CSVコンバーター"
    OPEN_FILE_DIALOG = "CSVファイルを開く"
    CONFIRM_DIALOG = "確認"
    ERROR_DIALOG = "エラー"
    END_DIALOG = "変換完了"


class MessageText:
    MAIN_DIALOG = "変換したいCSVファイルを選択し、変換ボタンを押してください。"
    FILE = "ファイル:"
    CONFIRM_DIALOG = "選択したファイルを変換してもよろしいですか？"
    WRONG_EXT_DIALOG = "異なる形式のファイルが選択されました。\ncsv形式のファイルを選択し直してください。"
    BLANK_ENTRY_DIALOG = "ファイル名を指定してください。"
    NOT_EXIST_FILE_DIALOG = "存在しないファイルです。\nファイルを選択し直してください。"
    FILE_OPEN_ERROR_DIALOG = "ファイルの読み込みに失敗しました。"
    FILE_WRITE_ERROR_DIALOG = "ファイルの書き込みに失敗しました。"
    END_DIALOG = "ファイルの変換が完了しました。"


class ButtonText:
    OPEN = "参照"
    CONVERT = "変換"
    CLOSE = "閉じる"
    OK = "OK"
    YES = "はい"
    NO = "いいえ"


class LabelWidth:
    CONFIRM_DIALOG = 35
    WRONG_EXT_DIALOG = 35
    BLANK_ENTRY_DIALOG = 30
    NOT_EXIST_FILE_DIALOG = 30
    FILE_OPEN_ERROR_DIALOG = 30
    FILE_WRITE_ERROR_DIALOG = 30
    END_DIALOG = 30


def main():
    main_dialog()


def main_dialog():
    root = Tk()
    root.title(DialogTitle.MAIN_DIALOG)
    posx = int(root.winfo_screenwidth() / 3)
    posy = int(root.winfo_screenheight() / 3)
    root.geometry("+" + str(posx) + "+" + str(posy))

    frame1 = ttk.Frame(root, padding=10)

    label_description = ttk.Label(frame1, text=MessageText.MAIN_DIALOG)

    frame2 = ttk.Frame(root, padding=10)

    lavel_file = ttk.Label(
        frame2,
        text=MessageText.FILE,
    )

    entry_text = StringVar()
    file_entry = ttk.Entry(frame2, textvariable=entry_text, width=30)

    button_file = ttk.Button(
        frame2, text=ButtonText.OPEN, command=lambda: file_button_clicked(entry_text)
    )

    frame3 = ttk.Frame(root, padding=10)

    button_convert = ttk.Button(
        frame3, text=ButtonText.CONVERT, command=lambda: entry_check(entry_text.get())
    )

    button_close = ttk.Button(
        frame3, text=ButtonText.CLOSE, command=lambda: [root.destroy(), exit()]
    )

    # ダイアログを表示
    frame1.grid()
    frame2.grid()
    frame3.grid()
    label_description.grid(row=0, column=0)
    lavel_file.grid(row=0, column=0)
    file_entry.grid(row=0, column=1)
    button_file.grid(row=0, column=2)
    button_convert.grid(row=0, column=0)
    button_close.grid(row=0, column=1)
    root.mainloop()


def file_button_clicked(entry_text):
    file_name = filedialog.askopenfilename(
        title=DialogTitle.OPEN_FILE_DIALOG,
        filetypes=[("CSVファイル", "*.csv")],
        initialdir=os.path.dirname(os.path.abspath(sys.executable)),
    )

    entry_text.set(file_name)


def entry_check(entry_text):
    if entry_text == "":
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.BLANK_ENTRY_DIALOG,
            LabelWidth.BLANK_ENTRY_DIALOG,
        )
        return

    if not os.path.isfile(entry_text):
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.NOT_EXIST_FILE_DIALOG,
            LabelWidth.NOT_EXIST_FILE_DIALOG,
        )
        return

    if not entry_text.endswith(".csv"):
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.WRONG_EXT_DIALOG,
            LabelWidth.WRONG_EXT_DIALOG,
        )
        return

    confirm_dialog(entry_text)


def confirm_dialog(entry_text):
    root = Tk()
    root.title(DialogTitle.CONFIRM_DIALOG)
    posx = int(root.winfo_screenwidth() / 5 * 2)
    posy = int(root.winfo_screenheight() / 5 * 2)
    root.geometry("+" + str(posx) + "+" + str(posy))

    frame1 = ttk.Frame(root, padding=10)

    label_message = ttk.Label(
        frame1, text=MessageText.CONFIRM_DIALOG, width=35, padding=(10)
    )

    frame2 = ttk.Frame(root, padding=10)

    button_yes = ttk.Button(
        frame2,
        text=ButtonText.YES,
        command=lambda: [root.destroy(), srpg_studio_csv_convert(entry_text)],
    )

    button_no = ttk.Button(frame2, text=ButtonText.NO, command=lambda: root.destroy())

    # ダイアログを表示
    frame1.grid()
    frame2.grid()
    label_message.grid(row=0, column=0)
    button_yes.grid(row=1, column=0)
    button_no.grid(row=1, column=1)
    root.mainloop()


def ok_button_dialog(dialog_title, message_text, label_width):
    root = Tk()
    root.title(dialog_title)
    posx = int(root.winfo_screenwidth() / 5 * 2)
    posy = int(root.winfo_screenheight() / 5 * 2)
    root.geometry("+" + str(posx) + "+" + str(posy))

    frame1 = ttk.Frame(root, padding=10)

    label_message = ttk.Label(
        frame1, text=message_text, width=label_width, padding=(10)
    )

    button_ok = ttk.Button(frame1, text=ButtonText.OK, command=lambda: root.destroy())

    # ダイアログを表示
    frame1.grid()
    label_message.grid(row=0, column=0)
    button_ok.grid(row=1, column=0)
    root.mainloop()


def is_half_width_digit(num_str):
    if num_str.isascii() and num_str.isdecimal():
        return True
    else:
        return False


def srpg_studio_csv_convert(entry_text):
    try:
        with open(entry_text, encoding="utf-8") as in_file:
            reader = csv.reader(in_file)
            row_list = [tmp_row for tmp_row in reader]
    except:
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.FILE_OPEN_ERROR_DIALOG,
            LabelWidth.FILE_OPEN_ERROR_DIALOG,
        )

    try:
        with open(entry_text[:-4] + ".txt", mode="w", encoding="utf-8") as out_file:
            # 0:形式 1:発言者 2:位置 3:表情 4:内容
            for row in row_list:
                if row[0] == "メッセージ":
                    if row[1] != "":
                        out_file.write("\n")
                        out_file.write(row[1] + "：" + row[2])
                        if row[3] != "":
                            out_file.write("：" + row[3] + "\n")
                        else:
                            out_file.write("\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "テロップ":
                    if row[2] != "":
                        out_file.write("\n")
                        out_file.write("テロップ：" + row[2] + "\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "メッセージタイトル":
                    if row[2] != "":
                        out_file.write("\n")
                        out_file.write("タイトル：" + row[2] + "\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "スチルメッセージ":
                    if row[1] != "":
                        out_file.write("\n")
                        out_file.write(row[1] + "：スチル\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "情報ウィンドウ":
                    if row[1] != "":
                        out_file.write("\n")
                        out_file.write("情報：" + row[1] + "\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "メッセージスクロール":
                    if row[2] != "":
                        out_file.write("\n")
                        out_file.write("スクロール：" + row[2] + "\n")
                        out_file.write(row[4] + "\n")
                    else:
                        out_file.write(row[4] + "\n")
                elif row[0] == "選択肢":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    choices = row[4].split(",")
                    out_file.write("\n")
                    out_file.write("選択肢：" + num + "\n")
                    for c in choices:
                        out_file.write(c + "\n")
                elif row[0] == "【場所イベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<PL" + num + ">\n")
                elif row[0] == "【自動開始イベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<AT" + num + ">\n")
                elif row[0] == "【会話イベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<TK" + num + ">\n")
                elif row[0] == "【オープニングイベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<OP" + num + ">\n")
                elif row[0] == "【エンディングイベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<ED" + num + ">\n")
                elif row[0] == "【コミュニケーションイベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<CM" + num + ">\n")
                elif row[0] == "【回想イベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<RE" + num + ">\n")
                elif row[0] == "【マップ共有イベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<MC" + num + ">\n")
                elif row[0] == "【ブックマークイベント】":
                    if is_half_width_digit(row[1]):
                        num = row[1]
                    else:
                        num = "0"
                    out_file.write("\n")
                    out_file.write("<BK" + num + ">\n")
    except:
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.FILE_WRITE_ERROR_DIALOG,
            LabelWidth.FILE_WRITE_ERROR_DIALOG,
        )

    ok_button_dialog(
        DialogTitle.END_DIALOG, MessageText.END_DIALOG, LabelWidth.END_DIALOG
    )


if __name__ == "__main__":
    main()
