import win32com.client
from datetime import datetime, timedelta
import pythoncom
import streamlit as st

def create_end_of_work_email(your_name, group, start_time_str, boss, cc_list):
    pythoncom.CoInitialize()

    today = datetime.now()
    today_str = f"{today.month}/{today.day}"
    subject = f"【勤務終了】 ({group}) {your_name} {today_str}"

    start_time = datetime.strptime(start_time_str, "%H:%M")
    end_time = (datetime.now() + timedelta(minutes=15)).strftime("%H:%M")
    work_time = f"{start_time_str}~{end_time}"

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    today_start = datetime(today.year, today.month, today.day)
    today_end = today_start + timedelta(days=1)
    restriction = "[Start] >= '" + today_start.strftime("%m/%d/%Y %H:%M %p") + "' AND [End] < '" + today_end.strftime("%m/%d/%Y %H:%M %p") + "'"
    today_items = items.Restrict(restriction)

    schedule_titles = []
    for item in today_items:
        try:
            schedule_titles.append(f"・{item.Subject}")
        except Exception:
            pass

    schedule_text = "\n".join(schedule_titles) if schedule_titles else "・本日の予定はありません"

    body = f"""

本日の業務を終了します

{schedule_text}

勤務時間：{work_time}
"""

    mail = outlook.CreateItem(0)
    mail.To = boss
    mail.CC = "; ".join(cc_list)
    mail.Subject = subject
    mail.Body = body
    mail.Display()

# StreamlitのUI
st.title("勤務終了メール作成ツール")

# 入力フォーム
your_name = st.text_input("あなたの名前", "山田")
group = st.text_input("グループ名", "知財1GR3")
start_time_str = st.text_input("業務開始時間 (例: 8:45)", "8:45")
boss = st.text_input("上司のメールアドレス", "ishida@example.com")
cc_list = st.text_input("CCに送る人 (カンマ区切りで)", "mizutani@example.com, murakoshi@example.com")

# メール作成ボタン
if st.button("勤務終了メールを作成"):
    create_end_of_work_email(your_name, group, start_time_str, boss, cc_list)
    st.success("勤務終了メールが作成されました！")
