
import PySimpleGUI as sg
import pandas as pd
from make_sheets import main
import popup_text
import os


desktop_dir = os.path.expanduser('~/Desktop')


# GUIのレイアウト
layout = [[sg.Text("遅延レポートの作成"), sg.Button("選択するデータの説明", button_color=(), pad=(200, 0))],
          [sg.HorizontalSeparator(color="#778899")],
          [sg.Text("")],
          [sg.Text('請求データのファイルを選択してください'),  sg.FileBrowse(
              key="-FILE_BILLING-", target="-BILLING-")],
          [sg.Text('選択されたデータ:'), sg.Input(
              key='-BILLING-', readonly=True, size=(80, 1))],
          [sg.Text('入金データのファイルを選択してください'),  sg.FileBrowse(
              key="-FILE_PAYMENT-", target="-PAYMENT-")],
          [sg.Text('選択されたデータ:'), sg.Input(
              key='-PAYMENT-', readonly=True, size=(80, 1))],
          [sg.Text("")],
          [sg.Text('集計期間を選択してください:')],
          [sg.Text('開始日:'), sg.CalendarButton('開始日', target='-STARTDATE-', key='-START-', format='%Y-%m-%d'),
           sg.Input(key='-STARTDATE-', readonly=True, default_text="1900-01-01")],
          [sg.Text('終了日:'), sg.CalendarButton('終了日', target='-ENDDATE-', key='-END-', format='%Y-%m-%d'),
           sg.Input(key='-ENDDATE-', readonly=True, default_text="2103-03-27")],
          [sg.Text("")],
          [sg.Button('遅延レポート出力'), sg.Button('遅延レポートについての説明', button_color=(), pad=(200, 0))]]


sg.theme('SystemDefault1')
# GUIのウィンドウを作成
window = sg.Window('秋田屋入金請求データ分析', layout, resizable=True,
                   grab_anywhere=True, auto_size_buttons=True)

# データフレームが作成済みかどうかのフラグ
df_created = False

# GUIのイベントループ
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == '終了':
        break
    elif event == '選択するデータの説明':
        sg.popup(popup_text.text1, title='選択するデータの説明')
    elif event == '遅延レポート出力':
        if values["-PAYMENT-"] == "" or values["-BILLING-"] == "":
            sg.popup("ファイルが選択されていません。'Browse'ボタンからファイルを選択して下さい", title="")
            continue
        sg.popup_no_buttons("しばらくお待ちください", title="",
                            auto_close=True, auto_close_duration=0.5)
        start = values["-STARTDATE-"]
        end = values["-ENDDATE-"]
        start_date = pd.to_datetime(start, format="%Y-%m-%d")
        end_date = pd.to_datetime(end, format="%Y-%m-%d")
        main_dict = main(
            file_path_billing=values["-BILLING-"], file_path_payment=values["-PAYMENT-"], start_date=start_date, end_date=end_date)
        if main_dict == "error":
            continue

        df_delay = main_dict["df_delay"]
        df_delay_detail = main_dict["df_delay_detail"]
        everyday_df = main_dict["everyday_df"]
        everymonth_df = main_dict["everymonth_df"]
        main_df = main_dict["main_df"]

        with pd.ExcelWriter(f"{desktop_dir}/遅延レポート.xlsx") as writer:
            df_delay.to_excel(writer, sheet_name='遅延回数')
            df_delay_detail.to_excel(writer, sheet_name='遅延回数_詳細', index=False)
            everyday_df.to_excel(writer, sheet_name='日毎の集計')
            everymonth_df.to_excel(writer, sheet_name='月毎の集計')
            main_df.to_excel(writer, sheet_name='参照元データ', index=False)


        df_created = True
        sg.popup('遅延レポートを作成し、"遅延レポート.xlsx"のファイル名でデスクトップに出力しました。', title="")

    elif event == '遅延レポートについての説明':
        sg.popup(popup_text.text2, title='遅延レポートについての説明')

