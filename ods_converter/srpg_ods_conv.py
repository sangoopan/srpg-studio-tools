# Copyright (c) 2022 sangoopan
# Released under the MIT license.
# https://licenses.opensource.jp/MIT/MIT.html

import os
import sys
from tkinter import Frame, StringVar, Tk, filedialog, ttk

import pandas as pd
from tkinterdnd2 import *


class DialogTitle:
    MAIN_DIALOG = "SRPG Studio用 ODSコンバーター"
    OPEN_FILE_DIALOG = "ODSファイルを開く"
    ERROR_DIALOG = "エラー"
    END_DIALOG = "変換完了"


class MessageText:
    FILE_SELECT = "変換したいODSファイルを選択し、[読み込む]ボタンを押してください。"
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


class SelectedMethodValue:
    ALL_SHEET_CONVERT = "all"
    ONE_SHEET_CONVERT = "one"


class Application(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)

        # 親ウィンドウの設定
        self.title(DialogTitle.MAIN_DIALOG)
        posx = int(self.winfo_screenwidth() / 3)
        posy = int(self.winfo_screenheight() / 3)
        self.geometry(f"330x120+{posx}+{posy}")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ファイル選択フレームの設定
        self.file_select_frame = Frame()
        self.file_select_frame.drop_target_register(DND_FILES)
        self.file_select_frame.dnd_bind("<<Drop>>", func=self.drop_file)

        # 1段目
        self.label_discription = ttk.Label(
            self.file_select_frame, text=MessageText.FILE_SELECT
        )

        # 2段目
        self.label_file = ttk.Label(self.file_select_frame, text=MessageText.FILE)
        self.file_entry_text = StringVar()
        self.file_entry = ttk.Entry(
            self.file_select_frame, textvariable=self.file_entry_text
        )
        self.button_file = ttk.Button(
            self.file_select_frame,
            text=ButtonText.OPEN,
            command=self.file_button_clicked,
        )

        # 3段目
        self.button_read = ttk.Button(
            self.file_select_frame,
            text=ButtonText.READ,
            command=self.read_button_clicked,
        )
        self.button_close = ttk.Button(
            self.file_select_frame,
            text=ButtonText.CLOSE,
            command=self.close_button_clicked,
        )

        # ファイル選択フレームの配置
        self.file_select_frame.grid(row=0, column=0, sticky="NSEW")
        self.file_select_frame.grid_rowconfigure(1, weight=1)
        self.file_select_frame.grid_columnconfigure(1, weight=1)
        self.file_select_frame.grid_columnconfigure(2, weight=1)
        self.label_discription.grid(row=0, column=0, columnspan=4, pady=12)
        self.label_file.grid(row=1, column=0)
        self.file_entry.grid(row=1, column=1, columnspan=2, sticky="EW")
        self.button_file.grid(row=1, column=3, ipadx=2)
        self.button_read.grid(row=2, column=2, pady=12, sticky="E")
        self.button_close.grid(row=2, column=3, pady=12, sticky="E")

        # 変換方法選択フレームの設定
        self.method_select_frame = Frame()
        self.df_dict = dict()
        self.sheet_names = list()

        # 1段目
        self.selected_method = StringVar(value=SelectedMethodValue.ONE_SHEET_CONVERT)
        self.radio_all_sheet = ttk.Radiobutton(
            self.method_select_frame,
            text=MessageText.ALL_SHEET_CONVERT,
            value=SelectedMethodValue.ALL_SHEET_CONVERT,
            variable=self.selected_method,
        )
        self.radio_all_sheet.bind("<ButtonPress>", self.method_radio_clicked)
        self.radio_one_sheet = ttk.Radiobutton(
            self.method_select_frame,
            text=MessageText.ONE_SHEET_CONVERT,
            value=SelectedMethodValue.ONE_SHEET_CONVERT,
            variable=self.selected_method,
        )
        self.radio_one_sheet.bind("<ButtonPress>", self.method_radio_clicked)

        # 2段目
        self.label_selected_sheet = ttk.Label(
            self.method_select_frame,
            text=MessageText.SELECTED_SHEET,
        )
        self.selected_sheet_name = StringVar()
        self.combo_sheets = ttk.Combobox(
            self.method_select_frame, values=self.sheet_names, state="readonly"
        )
        self.combo_sheets.bind("<<ComboboxSelected>>", self.combo_clicked)

        # 3段目
        self.button_convert = ttk.Button(
            self.method_select_frame,
            text=ButtonText.CONVERT,
            command=self.convert_button_clicked,
        )
        self.button_cancel = ttk.Button(
            self.method_select_frame,
            text=ButtonText.CANCEL,
            command=self.cancel_button_clicked,
        )

        # 変換方法選択フレームの配置
        self.method_select_frame.grid(row=0, column=0, sticky="NSEW")
        self.method_select_frame.grid_rowconfigure(2, weight=1)
        self.method_select_frame.grid_columnconfigure(1, weight=1)
        self.radio_all_sheet.grid(row=0, column=0, sticky="W")
        self.radio_one_sheet.grid(row=1, column=0, sticky="W")
        self.label_selected_sheet.grid(row=2, column=0, sticky="E")
        self.combo_sheets.grid(row=2, column=1, columnspan=5, sticky="EW")
        self.button_convert.grid(row=3, column=4, pady=12, sticky="E")
        self.button_cancel.grid(row=3, column=5, pady=12, sticky="E")

        # ファイル選択フレームを前面に
        self.file_select_frame.tkraise()

    # [参照]ボタンが押されたらファイル選択画面を開く
    def file_button_clicked(self):
        file_name = filedialog.askopenfilename(
            title=DialogTitle.OPEN_FILE_DIALOG,
            filetypes=[("ODSファイル", "*.ods")],
            initialdir=os.path.dirname(os.path.abspath(sys.executable)),
        )

        self.file_entry_text.set(file_name)

    # ドラッグ&ドロップしたファイルのパスを取得
    # ファイルが複数のときは先頭のパスのみ
    def drop_file(self, event=None):
        if event:
            file_paths = self.tk.splitlist(event.data)
            self.file_entry_text.set(file_paths[0])

    # [読み込む]ボタンが押されたらファイルのパスと存在チェック
    # 問題なければファイルを読み込み、変換方法選択フレームに遷移
    def read_button_clicked(self):
        entry_text = self.file_entry_text.get()
        if entry_text == "":
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.BLANK_ENTRY_DIALOG
            )
            dialog.mainloop()
            return
        elif not os.path.isfile(entry_text):
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.NOT_EXIST_FILE_DIALOG
            )
            dialog.mainloop()
            return
        elif not entry_text.endswith(".ods"):
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.WRONG_EXT_DIALOG
            )
            dialog.mainloop()
            return

        try:
            self.df_dict = pd.read_excel(entry_text, sheet_name=None, engine="odf")
        except:
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.FILE_OPEN_ERROR_DIALOG
            )
            dialog.mainloop()
        self.sheet_names = list(self.df_dict.keys())
        self.selected_sheet_name.set(self.sheet_names[0])
        self.combo_sheets["values"] = self.sheet_names
        self.combo_sheets.set(self.sheet_names[0])
        self.method_select_frame.tkraise()

    # [閉じる]ボタンが押されたらウィジェット削除
    def close_button_clicked(self):
        self.destroy()

    # ラジオボタンが押されたらコンボボックスの活性/非活性を切り替える
    def method_radio_clicked(self, event):
        if event.widget.cget("value") == SelectedMethodValue.ALL_SHEET_CONVERT:
            self.combo_sheets["state"] = "disable"
        else:
            self.combo_sheets["state"] = "readonly"

    # コンボボックスから選択したシート名で更新
    def combo_clicked(self, event):
        self.selected_sheet_name.set(self.combo_sheets.get())

    # [変換]ボタンが押されたらラジオボタンの状態に応じて変換処理
    def convert_button_clicked(self):
        if self.selected_method.get() == SelectedMethodValue.ALL_SHEET_CONVERT:
            self.all_sheet_convert()
        else:
            self.one_sheet_convert()

    # [キャンセル]ボタンが押されたらファイル選択フレームに遷移
    def cancel_button_clicked(self):
        self.file_select_frame.tkraise()

    # 1シート変換
    def one_sheet_convert(self):
        entry_text = self.file_entry_text.get()
        sheet_name = self.selected_sheet_name.get()
        df = self.df_dict[sheet_name]
        df.fillna("", inplace=True)
        row_list = df.astype(str).values.tolist()
        out_text_list = []
        for row in row_list:
            self.current_line_convert(row, out_text_list)

        try:
            with open(
                f"{entry_text[:-4]}_{sheet_name}.txt",
                mode="w",
                encoding="utf-8",
            ) as out_file:
                out_file.write("".join(out_text_list))
        except:
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.FILE_WRITE_ERROR_DIALOG
            )
            dialog.mainloop()
            return

        dialog = OneButtonDialog(Tk(), DialogTitle.END_DIALOG, MessageText.END_DIALOG)
        dialog.mainloop()

    # 全シート変換
    def all_sheet_convert(self):
        entry_text = self.file_entry_text.get()
        out_text_list = []
        for df in self.df_dict.values():
            df.fillna("", inplace=True)
            row_list = df.astype(str).values.tolist()
            for row in row_list:
                self.current_line_convert(row, out_text_list)
            out_text_list.append("\n")

        try:
            with open(f"{entry_text[:-4]}.txt", mode="w", encoding="utf-8") as out_file:
                out_file.write("".join(out_text_list))
        except:
            dialog = OneButtonDialog(
                Tk(), DialogTitle.ERROR_DIALOG, MessageText.FILE_WRITE_ERROR_DIALOG
            )
            dialog.mainloop()
            return

        dialog = OneButtonDialog(Tk(), DialogTitle.END_DIALOG, MessageText.END_DIALOG)
        dialog.mainloop()

    # フォーマットに沿った変換処理
    def current_line_convert(self, row, out_text_list):
        # 0:形式 1:id 2:名前 3:位置 4:表情 5:内容
        if row[0] == "メッセージ":
            if row[2] != "":
                out_text_list.append(f"\n{row[2]}：{row[3]}")
                if row[4] != "":
                    out_text_list.append(f"：{row[4]}\n")
                else:
                    out_text_list.append("\n")
                out_text_list.append(f"{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "テロップ":
            if row[3] != "":
                out_text_list.append(f"\nテロップ：{row[3]}\n{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "メッセージタイトル":
            if row[3] != "":
                out_text_list.append(f"\nタイトル：{row[3]}\n{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "スチルメッセージ":
            if row[2] != "":
                out_text_list.append(f"\n{row[2]}：スチル\n{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "情報ウィンドウ":
            if row[2] != "":
                out_text_list.append(f"\n情報：{row[2]}\n{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "メッセージスクロール":
            if row[3] != "":
                out_text_list.append(f"\nスクロール：{row[3]}\n{row[5]}\n")
            else:
                out_text_list.append(f"{row[5]}\n")
        elif row[0] == "選択肢":
            out_text_list.append(f"\n選択肢：{self.convert_id(row[1])}\n")
            choices = row[5].split(",")
            for c in choices:
                out_text_list.append(f"{c}\n")
        elif row[0] == "【ユニットイベント】":
            out_text_list.append(f"\n<({self.convert_id(row[1])}){row[2]}>\n")
        elif row[0] == "【場所イベント】":
            out_text_list.append(f"\n<PL{self.convert_id(row[1])}>\n")
        elif row[0] == "【自動開始イベント】":
            out_text_list.append(f"\n<AT{self.convert_id(row[1])}>\n")
        elif row[0] == "【会話イベント】":
            out_text_list.append(f"\n<TK{self.convert_id(row[1])}>\n")
        elif row[0] == "【オープニングイベント】":
            out_text_list.append(f"\n<OP{self.convert_id(row[1])}>\n")
        elif row[0] == "【エンディングイベント】":
            out_text_list.append(f"\n<ED{self.convert_id(row[1])}>\n")
        elif row[0] == "【コミュニケーションイベント】":
            out_text_list.append(f"\n<CM{self.convert_id(row[1])}>\n")
        elif row[0] == "【回想イベント】":
            out_text_list.append(f"\n<RE{self.convert_id(row[1])}>\n")
        elif row[0] == "【マップ共有イベント】":
            out_text_list.append(f"\n<MC{self.convert_id(row[1])}>\n")
        elif row[0] == "【ブックマークイベント】":
            out_text_list.append(f"\n<BK{self.convert_id(row[1])}>\n")

    # id欄の文字列を整数値に変換
    def convert_id(self, num_str):
        try:
            float(num_str)
        except ValueError:
            return 0
        else:
            return int(float(num_str))


class OneButtonDialog(Frame):
    def __init__(self, root: Tk, dialog_title, message_text):
        super().__init__(root)
        root.title(dialog_title)
        posx = int(root.winfo_screenwidth() / 5 * 2)
        posy = int(root.winfo_screenheight() / 5 * 2)
        root.geometry(f"220x95+{posx}+{posy}")
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # 1段目
        self.label_message = ttk.Label(self, text=message_text)

        # 2段目
        self.button_ok = ttk.Button(self, text=ButtonText.OK, command=root.destroy)

        # ウィジェットの配置
        self.grid(row=0, column=0, sticky="NSEW")
        self.label_message.pack(pady=10)
        self.button_ok.pack(pady=5)
