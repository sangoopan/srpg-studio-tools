# Copyright (c) 2022 sangoopan
# Released under the MIT license
# https://licenses.opensource.jp/MIT/MIT.html

import os
import sys
from tkinter import EW, Frame, StringVar, Tk, filedialog, ttk

import pandas as pd


class DialogTitle:
    MAIN_DIALOG = "SRPG Studio用 ODSコンバーター"
    OPEN_FILE_DIALOG = "ODSファイルを開く"
    METHOD_SELECT_DIALOG = "変換方法の選択"
    ERROR_DIALOG = "エラー"
    END_DIALOG = "変換完了"


class MessageText:
    MAIN_DIALOG = "変換したいODSファイルを選択し、[読み込む]ボタンを押してください。"
    FILE = "ファイル:"
    ALL_SHEET_CONVERT = "全てのシートを変換"
    ONE_SHEET_CONVERT = "1つのシートを変換"
    SELECTED_SHEET = "変換するシート:"
    WRONG_EXT_DIALOG = "異なる形式のファイルが選択されました。\nods形式のファイルを選択し直してください。"
    BLANK_ENTRY_DIALOG = "ファイル名を指定してください。"
    NOT_EXIST_FILE_DIALOG = "存在しないファイルです。\nファイルを選択し直してください。"
    FILE_OPEN_ERROR_DIALOG = "ファイルの読み込みに失敗しました。"
    FILE_WRITE_ERROR_DIALOG = "ファイルの書き込みに失敗しました。"
    END_DIALOG = "ファイルの変換が完了しました。"


class ButtonText:
    OPEN = "参照"
    READ = "読み込む"
    CONVERT = "変換"
    CLOSE = "閉じる"
    OK = "OK"
    CANCEL = "キャンセル"
    YES = "はい"
    NO = "いいえ"


class LabelWidth:
    WRONG_EXT_DIALOG = 35
    BLANK_ENTRY_DIALOG = 30
    NOT_EXIST_FILE_DIALOG = 30
    FILE_OPEN_ERROR_DIALOG = 30
    FILE_WRITE_ERROR_DIALOG = 30
    END_DIALOG = 30


class SelectedMethodValue:
    ALL_SHEET_CONVERT = "all"
    ONE_SHEET_CONVERT = "one"


class MethodSelectDialog(Frame):
    def __init__(self, root=None, entry_text=""):
        self.entry_text = entry_text
        try:
            self.df_dict = pd.read_excel(self.entry_text, sheet_name=None, engine="odf")
            self.sheet_names = list(self.df_dict.keys())
        except:
            ok_button_dialog(
                DialogTitle.ERROR_DIALOG,
                MessageText.FILE_OPEN_ERROR_DIALOG,
                LabelWidth.FILE_OPEN_ERROR_DIALOG,
            )

        super().__init__(root)
        root.title(DialogTitle.METHOD_SELECT_DIALOG)
        posx = int(root.winfo_screenwidth() / 5 * 2)
        posy = int(root.winfo_screenheight() / 5 * 2)
        root.geometry("300x135+" + str(posx) + "+" + str(posy))

        self.selected_method = StringVar()
        self.selected_method.set(value=SelectedMethodValue.ONE_SHEET_CONVERT)
        self.radio_all_sheet = ttk.Radiobutton(
            self,
            text=MessageText.ALL_SHEET_CONVERT,
            value=SelectedMethodValue.ALL_SHEET_CONVERT,
            variable=self.selected_method,
        )
        self.radio_one_sheet = ttk.Radiobutton(
            self,
            text=MessageText.ONE_SHEET_CONVERT,
            value=SelectedMethodValue.ONE_SHEET_CONVERT,
            variable=self.selected_method,
        )
        self.radio_all_sheet.bind("<ButtonPress>", self.select_radio)
        self.radio_one_sheet.bind("<ButtonPress>", self.select_radio)

        self.label_selected_sheet = ttk.Label(
            self,
            text=MessageText.SELECTED_SHEET,
            width=12,
            padding=(10),
        )

        self.selected_sheet_name = StringVar()
        self.combo_sheets = ttk.Combobox(
            self,
            textvariable=self.selected_sheet_name,
            values=self.sheet_names,
            state="readonly",
            width=20,
        )
        self.selected_sheet_name.set(self.sheet_names[0])
        self.combo_sheets.set(self.sheet_names[0])
        self.combo_sheets.bind("<<ComboboxSelected>>", self.select_combo)

        self.button_convert = ttk.Button(
            self,
            text=ButtonText.CONVERT,
            command=lambda: [root.destroy(), self.convert_fork()],
        )
        self.button_cancel = ttk.Button(
            self, text=ButtonText.CANCEL, command=root.destroy
        )

        self.grid()
        self.radio_all_sheet.grid(row=0, column=0, padx=10, pady=8)
        self.radio_one_sheet.grid(row=0, column=1, padx=10, pady=8)
        self.label_selected_sheet.grid(row=1, column=0, padx=10, pady=8, sticky=EW)
        self.combo_sheets.grid(row=1, column=1, padx=10, pady=8, sticky=EW)
        self.button_convert.grid(row=2, column=0, padx=5, pady=9)
        self.button_cancel.grid(row=2, column=1, padx=5, pady=9)
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)

    def select_radio(self, event):
        if event.widget.cget("value") == SelectedMethodValue.ALL_SHEET_CONVERT:
            self.selected_method.set(SelectedMethodValue.ALL_SHEET_CONVERT)
            self.combo_sheets["state"] = "disable"
        else:
            self.selected_method.set(SelectedMethodValue.ONE_SHEET_CONVERT)
            self.combo_sheets["state"] = "readonly"

    def select_combo(self, event):
        self.selected_sheet_name.set(self.combo_sheets.get())

    def convert_fork(self):
        if self.selected_method.get() == SelectedMethodValue.ALL_SHEET_CONVERT:
            all_sheet_convert(self.entry_text, self.df_dict)
        else:
            sheet_name = self.selected_sheet_name.get()
            df = self.df_dict[sheet_name]
            one_sheet_convert(self.entry_text, df, sheet_name)


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

    button_read = ttk.Button(
        frame3, text=ButtonText.READ, command=lambda: entry_check(entry_text.get())
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
    button_read.grid(row=0, column=0, padx=15)
    button_close.grid(row=0, column=1, padx=15)
    root.mainloop()


def file_button_clicked(entry_text):
    file_name = filedialog.askopenfilename(
        title=DialogTitle.OPEN_FILE_DIALOG,
        filetypes=[("ODSファイル", "*.ods")],
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

    if not entry_text.endswith(".ods"):
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.WRONG_EXT_DIALOG,
            LabelWidth.WRONG_EXT_DIALOG,
        )
        return

    method_select_dialog(entry_text)


def method_select_dialog(entry_text):
    root = Tk()
    dialog = MethodSelectDialog(root, entry_text)
    dialog.mainloop()


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


def one_sheet_convert(entry_text, df, sheet_name):
    try:
        with open(
            entry_text[:-4] + "_" + sheet_name + ".txt", mode="w", encoding="utf-8"
        ) as out_file:
            df.fillna("", inplace=True)
            for i in range(0, len(df)):
                sr = df.iloc[i]
                row = [str(elem) for elem in sr.values.tolist()]
                current_line_convert(row, out_file)
    except:
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.FILE_WRITE_ERROR_DIALOG,
            LabelWidth.FILE_WRITE_ERROR_DIALOG,
        )

    ok_button_dialog(
        DialogTitle.END_DIALOG, MessageText.END_DIALOG, LabelWidth.END_DIALOG
    )


def all_sheet_convert(entry_text, df_dict):
    try:
        with open(entry_text[:-4] + ".txt", mode="w", encoding="utf-8") as out_file:
            for df in df_dict.values():
                df.fillna("", inplace=True)
                for i in range(0, len(df)):
                    sr = df.iloc[i]
                    row = [str(elem) for elem in sr.values.tolist()]
                    current_line_convert(row, out_file)
                out_file.write("\n")
    except:
        ok_button_dialog(
            DialogTitle.ERROR_DIALOG,
            MessageText.FILE_WRITE_ERROR_DIALOG,
            LabelWidth.FILE_WRITE_ERROR_DIALOG,
        )

    ok_button_dialog(
        DialogTitle.END_DIALOG, MessageText.END_DIALOG, LabelWidth.END_DIALOG
    )


def current_line_convert(row, out_file):
    # 0:形式 1:発言者 2:位置 3:表情 4:内容
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
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        choices = row[4].split(",")
        out_file.write("\n")
        out_file.write("選択肢：" + num + "\n")
        for c in choices:
            out_file.write(c + "\n")
    elif row[0] == "【場所イベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<PL" + num + ">\n")
    elif row[0] == "【自動開始イベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<AT" + num + ">\n")
    elif row[0] == "【会話イベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<TK" + num + ">\n")
    elif row[0] == "【オープニングイベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<OP" + num + ">\n")
    elif row[0] == "【エンディングイベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<ED" + num + ">\n")
    elif row[0] == "【コミュニケーションイベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<CM" + num + ">\n")
    elif row[0] == "【回想イベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<RE" + num + ">\n")
    elif row[0] == "【マップ共有イベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<MC" + num + ">\n")
    elif row[0] == "【ブックマークイベント】":
        num = "0"
        if is_half_width_digit(row[1]):
            num = row[1]
        out_file.write("\n")
        out_file.write("<BK" + num + ">\n")


if __name__ == "__main__":
    main()
