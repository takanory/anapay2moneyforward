"""
ANA Payの情報をメールから取得してスプレッドシートに書き込む

それからスプレッドシートの情報を元に、Money Fowardに情報を書き込む
"""

import base64
import logging
import os
from dataclasses import dataclass
from datetime import datetime

import gspread
import helium
from dateutil import parser
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Google Spreadsheet ID and Sheet name
SHEET_ID = "143Ewai1jFlt4d4msZI8fXersf2IErrzTQfFjjrwzOwM"
SHEET_NAME = "ANAPay"

MF_URL = "https://ssnb.x.moneyforward.com/cf"

format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
logging.basicConfig(format=format, level=logging.INFO)

load_dotenv()


@dataclass
class ANAPay:
    """ANA Pay information"""

    email_date: datetime = None
    date_of_use: datetime = None
    amount: int = 0
    store: str = ""

    def values(self) -> tuple[str, str, str, str]:
        """return tuple of values for spreadsheet"""
        return self.email_date_str, self.date_of_use_str, self.amount, self.store

    @property
    def email_date_str(self) -> str:
        return f"{self.email_date:%Y-%m-%d %H:%M:%S}"

    @property
    def date_of_use_str(self) -> str:
        return f"{self.date_of_use:%Y-%m-%d %H:%M:%S}"


def get_mail_info(res: dict) -> ANAPay | None:
    """
    1件のメールからANA Payの利用情報を取得して返す
    """
    ana_pay = ANAPay()
    for header in res["payload"]["headers"]:
        if header["name"] == "Date":
            date_str = header["value"].replace(" +0900 (JST)", "")
            ana_pay.email_date = parser.parse(date_str)

    # 本文から日時、金額、店舗を取り出す
    # ご利用日時：2023-06-28 22:46:19
    # ご利用金額：44,308円
    # ご利用店舗：SMOKEBEERFACTORY OTSUKATE
    data = res["payload"]["body"]["data"]
    body = base64.urlsafe_b64decode(data).decode()
    for line in body.splitlines():
        if line.startswith("ご利用"):
            key, value = line.split("：")
            if key == "ご利用日時":
                ana_pay.date_of_use = parser.parse(value)
            elif key == "ご利用金額":
                ana_pay.amount = int(value.replace(",", "").replace("円", ""))
            elif key == "ご利用店舗":
                ana_pay.store = value
    return ana_pay


def get_anapay_info(after: str) -> list[ANAPay]:
    """
    gmailからANA Payの利用履歴を取得する
    """
    ana_pay_list = []

    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    service = build("gmail", "v1", credentials=creds)

    # https://developers.google.com/gmail/api/reference/rest/v1/users.messages/list
    query = f"from:payinfo@121.ana.co.jp subject:ご利用のお知らせ after:{after}"
    results = service.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])
    for message in reversed(messages):
        # https://developers.google.com/gmail/api/reference/rest/v1/users.messages/get
        res = service.users().messages().get(userId="me", id=message["id"]).execute()
        ana_pay = get_mail_info(res)
        if ana_pay:
            ana_pay_list.append(ana_pay)
    return ana_pay_list

    after = "2023/06/28"


def get_last_email_date(records: list[dict[str, str]]):
    """get last email date for gmail search"""
    after = "2023/06/28"
    if records:
        last_email_date = parser.parse(records[-1]["email_date"])
        after = f"{last_email_date:%Y/%m/%d}"
    return after


def gmail2spredsheet(worksheet):
    """gmailからANA Payの利用履歴を取得しスプレッドシートに書き込む"""
    # get all records from spreadsheet
    records = worksheet.get_all_records()
    logging.info("Records in spreadsheet: %d", len(records))

    # get last day from records
    after = get_last_email_date(records)
    logging.info("Last day on spreadsheet: %s", after)
    email_date_set = set(parser.parse(r["email_date"]) for r in records)

    # get ANA Pay email from Gamil
    ana_pay_list = get_anapay_info(after)
    logging.info("ANA Pay emails: %d", len(ana_pay_list))

    # add ANA Pay record to spreadsheet
    count = 0
    for ana_pay in ana_pay_list:
        # メールの日付が存在しない場合はレコードを追加
        if ana_pay.email_date not in email_date_set:
            worksheet.append_row(ana_pay.values(), value_input_option="USER_ENTERED")
            count += 1
            logging.info("Record added to spreadsheet: %s", ana_pay.values())
    logging.info("Records added to spreadsheet: %d", count)


def login_mf():
    """login moneyforward sbi"""

    email = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")

    # https://selenium-python-helium.readthedocs.io/en/latest/api.html
    logging.info("Login to moneyfoward")
    helium.start_firefox(MF_URL)
    helium.write(email, into="メールアドレス")
    helium.write(password, into="パスワード")
    helium.click("ログイン")
    helium.wait_until(helium.Button("手入力").exists)


def add_mf_record(dt: datetime, amount: int, store: str, store_info: dict | None):
    """
    add record to moneyfoward
    """

    # https://selenium-python-helium.readthedocs.io/en/latest/api.html
    helium.click("手入力")
    # breakpoint()
    helium.write(f"{dt:%Y/%m/%d}", into="日付")
    helium.click("日付")

    helium.write(amount, into="支出金額")
    asset = helium.find_all(helium.ComboBox())[0]
    helium.select(asset, asset.options[2])

    if store_info:
        category = helium.find_all(helium.Link("未分類"))[0]
        l_category = helium.find_all(helium.S("#js-large-category-selected"))[0]
        helium.click(l_category)
        helium.click(store_info["大項目"])

        m_category = helium.find_all(helium.S("#js-middle-category-selected"))[0]
        helium.click(m_category)
        helium.click(store_info["中項目"])

        helium.write(store_info["店名"], into="内容をご入力下さい(任意)")
    else:
        helium.write(store, into="内容をご入力下さい(任意)")

    helium.click("保存する")
    logging.info(f"Record added to moneyforward: {dt:%Y/%m/%d}, {amount}, {store}")

    helium.wait_until(helium.Button("続けて入力する").exists)
    helium.click("続けて入力する")


def spreadsheet2mf(worksheet, store_dict: dict[str, dict[str, str]]) -> None:
    """スプレッドシートからmoneyfowardに書き込む"""

    records = worksheet.get_all_records()

    # すべてmoneyforwardに登録済みならなにもしない
    if all(record["mf"] == "done" for record in records):
        return

    login_mf()  # login to moneyfoward
    added = 0
    for count, record in enumerate(records):
        if record["mf"] != "done":
            date_of_use = parser.parse(record["date_of_use"])
            amount = int(record["amount"])
            store = record["store"]
            add_mf_record(date_of_use, amount, store, store_dict.get(store))

            # update spread sheets for "done" message
            worksheet.update_cell(count + 2, 5, "done")
            added += 1
    helium.kill_browser()

    logging.info(f"Records added to moneyforward: {added}")


def main():
    gc = gspread.oauth(
        credentials_filename="credentials.json", authorized_user_filename="token.json"
    )
    sheet = gc.open_by_key(SHEET_ID)
    anapay_sheet = sheet.worksheet("ANAPay")
    store_sheet = sheet.worksheet("ANAPayStore")
    store_dict = {store["store"]: store for store in store_sheet.get_all_records()}

    gmail2spredsheet(anapay_sheet)
    spreadsheet2mf(anapay_sheet, store_dict)


if __name__ == "__main__":
    main()
