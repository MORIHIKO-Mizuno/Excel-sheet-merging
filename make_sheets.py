

import pandas as pd
import PySimpleGUI as sg


# note: 必要なデータの入った列があるか確認する関数
def columns_check(df, columns_list):
    non_exist_list = []
    for col in columns_list:
        if col in df.columns:
            # note: 処理をスキップ
            continue
        else:
            # note: なかったものについてはnon_exist_listに加える
            non_exist_list.append(col)
    if len(non_exist_list) == 0:
        return None
    else:
        return non_exist_list

# NOTE: 二つのデータフレームを統合して大元のデータを作る、データ型を日付などに直す。
def make_main_df(df_billing, df_payment,start_date,end_date):

    # note: データの加工
    df_payment["商大種別"] = df_payment["商大種別"].astype("int")
    df_billing["identification"] = "df_billing"
    df_payment["identification"] = "df_payment"
    df_billing["回収予定日_入金日"] = df_billing["回収予定日"]
    df_payment["回収予定日_入金日"] = df_payment["計上日"]
    # note: 結合させるための下準備
    rename_dict = {"取引先コード": "得意先コード", "取引先名称": "取引先名称漢字"}
    df_billing = df_billing.rename(columns=rename_dict)
    main_df = pd.concat([df_billing, df_payment], join="outer")
    # note: 日付のデータ方に変更
    date_list = ["回収予定日", "計上日", "回収予定日_入金日"]
    main_df = main_df.replace({"回収予定日": {
        "0": "00000000"}, "計上日": {"0": "00000000"}, "回収予定日_入金日": {"0": "00000000"}})
    main_df[date_list] = main_df[date_list].apply(
        pd.to_datetime, format='%Y%m%d', errors="coerce").apply(lambda x: pd.to_datetime(x).round("D"))
    

    main_df = main_df.sort_values(
        ["得意先コード", "回収予定日_入金日", "商大種別", "identification"], ascending=[True, True, True, False])
    main_df = main_df[["部課コード",
                       "部課名称",
                       "得意先コード",
                       "取引先名称漢字",
                       "商大種別",
                       "商大種別名称",
                       "入金金額",
                       "差引繰越金額",
                       "差引売上金額",
                       "今回請求額",
                       "集金方法コード",
                       "集金方法",
                       "identification",
                       "回収予定日_入金日",
                       "金額"]]
    # main_df = main_df.loc[(main_df["回収予定日_入金日"]<= last) & (main_df["回収予定日_入金日"] >= first)]
    main_df.fillna({"今回請求額": 0, "金額": 0}, inplace=True)
    main_df=main_df.loc[(main_df["回収予定日_入金日"]>start_date)&(main_df["回収予定日_入金日"]<end_date)]
    return main_df


# Note：遅延期間、回数を調べる。df_delayを作る下処理
def check_delay(df):
    billing_payment = 0
    df.reset_index(drop=True, inplace=True)
    df = df.copy()
    # note: 収支列の作成
    for idx, billing, payment in zip(range(len(df)), df["今回請求額"], df["金額"]):
        billing_payment += billing+payment
        df.loc[idx, "収支"] = billing_payment
    # note: billingがない特殊ケースを除外
    if len(df.loc[df["identification"] == "df_billing", "回収予定日_入金日"]) == 0:
        df[["収支(修正)",	"遅延日数"]] = float("nan")
        return df
    idx = df.loc[df["identification"] == "df_billing", "回収予定日_入金日"].idxmin()
    col = df.columns.get_loc("収支")
    first = df.iloc[idx, col]
    adjustment = [-1*first if first > 0 else 0]
    df.loc[:, "収支(修正)"] = df["収支"].copy()+adjustment
    df.loc[:, "遅延日数"] = float('nan')

    # todo: 遅延のチェック

    check_required = df.loc[(df["identification"] ==
                             "df_billing") & (df["収支(修正)"] > 0)]
    for delay_billing_payment, day, delay_idx in zip(check_required["収支(修正)"], check_required["回収予定日_入金日"], check_required.index):
        payment_df = df.loc[(df["identification"] == "df_payment") & (
            df["回収予定日_入金日"] > day)]
        payment_df.reset_index(drop=True, inplace=True)
        idx = 0
        try:
            while delay_billing_payment > 0:
                delay_billing_payment += payment_df["金額"].iloc[idx]
                idx += 1
            else:
                payment_day = payment_df["回収予定日_入金日"].iloc[idx-1]
                delay = payment_day-day
                df.iloc[delay_idx, df.columns.get_loc("遅延日数")] = delay.days
        except IndexError:
            df.iloc[delay_idx, df.columns.get_loc("遅延日数")] = float('nan')
            continue

    return df

# Note：遅延回数を月ごとにまとめる。df_delayを作る下処理
def refining_df_delay(df_delay):
    df = pd.DataFrame(
        data=df_delay["遅延日数"].to_list(), index=df_delay["回収予定日_入金日"])
    df = df.resample("M").count()
    exceed_count_month = len(df[df.iloc[:, 0] > 0])
    average = df_delay["遅延日数"].mean()
    return exceed_count_month, average


# note; 遅延回数のシートのデータフレームを作成
def make_df_delay(main_df):
    df_delay = pd.DataFrame(None, columns=["得意先コード",'取引先名称漢字',"遅延回数", "遅延日数(平均)","部課コード", "部課名称",	"商大種別", "商大種別名称", "集金方法"])
    df_delay_detail=pd.DataFrame(None)
    main_list=main_df[["得意先コード",'取引先名称漢字']].drop_duplicates()
    df_delay[["得意先コード",'取引先名称漢字']]=main_list
    df_delay.set_index(["得意先コード",'取引先名称漢字'],inplace=True)
    print(df_delay)
    for company_code,company_name in zip(main_list["得意先コード"].tolist(),main_list['取引先名称漢字'].tolist()):
        
        df = main_df.loc[main_df["取引先名称漢字"] == company_name]
        check_delay_df=check_delay(df)
        df_delay_detail=pd.concat([check_delay_df,df_delay_detail],join="outer")
        exceed_count, average = refining_df_delay(check_delay_df)
        df_delay.loc[(company_code,company_name),["遅延回数", "遅延日数(平均)"]] = exceed_count, average

        for col in ["部課コード", "部課名称",	"商大種別", "商大種別名称", "集金方法"]:
            
            value=df[col].mode()
            
            if len(value)!=0:
                df_delay.loc[(company_code,company_name),col]=value.values[0]
    df_delay.sort_values("遅延回数",ascending=False,inplace=True)
    return {"df_delay":df_delay,"df_delay_detail":df_delay_detail}


# note: 集金方法をまとめる。日毎の集計の下処理
def make_method_list(main_df):
    value_counts = main_df['集金方法'].value_counts(normalize=False)
    sorted_counts = value_counts.sort_values(ascending=False)
    result_list = sorted_counts.index.tolist()
    return result_list

# note: 日毎の集計を作る
def make_everyday_df(main_df):
    method_list=make_method_list(main_df)

    everyday_df=main_df[["回収予定日_入金日","今回請求額","金額"]].groupby("回収予定日_入金日").sum()
    for method_ in method_list:
        everyday_df[method_]=main_df.loc[main_df["集金方法"]==method_,["回収予定日_入金日","今回請求額"]].groupby("回収予定日_入金日").sum()
        everyday_df.rename(columns={method_:f"請求額({method_})"},inplace=True)
    return everyday_df


# note: 全てを動かす関数。ファイルパスからファイルを読み込んでそれぞれの関数実行、またカラムチェック
def main(file_path_billing,file_path_payment,start_date,end_date):
    payment_columns_list = ["計上日",
                        "部課コード",
                        "部課名称",
                        "得意先コード",
                        "取引先名称漢字",
                        "商大種別",
                        "金額"]
    billing_columns_list = ["部課コード",
                        "部課名称",
                        "取引先コード",
                        "取引先名称",
                        "商大種別",
                        "入金金額",
                        "差引繰越金額",
                        "差引売上金額",
                        "今回請求額",
                        "回収予定日",
                        "集金方法コード",
                        "集金方法"]
    df_billing = pd.read_excel(file_path_billing, skiprows=1)
    result=columns_check(df_billing, billing_columns_list)
    if result:
        sg.popup(f'請求データのカラム名に\n{result}\nがありません。データの内容や空欄の個数などご確認ください')
        return 'error'

    df_payment = pd.read_excel(file_path_payment, skiprows=1)
    result=columns_check(df_payment, payment_columns_list)
    if result:
        sg.popup(f'入金データのカラム名に\n{result}\nがありません。データの内容や空欄の個数などご確認ください')
        return 'error'
    
    main_df=make_main_df(df_billing=df_billing,df_payment=df_payment,start_date=start_date,end_date=end_date)
    dict_=make_df_delay(main_df)
    df_delay=dict_["df_delay"]
    df_delay_detail=dict_["df_delay_detail"]
    everyday_df=make_everyday_df(main_df)
    everymonth_df=everyday_df.resample("M").sum()
    main_dict={"df_delay": df_delay, "df_delay_detail":df_delay_detail,"everyday_df": everyday_df, "everymonth_df":everymonth_df,"main_df": main_df}
    
    return main_dict



