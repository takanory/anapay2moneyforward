"""
ANA Payの情報をメールから取得してスプレッドシートに書き込む

それからスプレッドシートの情報を元に、Money Fowardに情報を書き込む
"""

import base64
from dataclasses import dataclass
from datetime import datetime

from dateutil import parser
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ご利用日時：2023-06-28 22:46:19
# ご利用金額：44,308円
# ご利用店舗：SMOKEBEERFACTORY OTSUKATE



@dataclass
class ANAPay:
    email_date: datetime = None
    subject: str = ""
    date_of_use: datetime = None
    amount: int = 0
    store: str = ""


def get_mail_info(service, mid):
    res = service.users().messages().get(userId="me", id=mid).execute()
    ana_pay = ANAPay()
    for header in res["payload"]["headers"]:
        if header["name"] == "Date":
            ana_pay.email_date = parser.parse(header["value"])
        elif header["name"] == "Subject":
            ana_pay.subject = header["value"]

    # ご利用のお知らせ以外は無視
    if "ご利用のお知らせ" not in ana_pay.subject:
        return

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


def get_anapay_mail(creds):
    service = build("gmail", "v1", credentials=creds)
    query = "from:payinfo@121.ana.co.jp"
    results = service.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])
    for message in messages[:5]:
        ana_pay = get_mail_info(service, message["id"])
        print(ana_pay)


def main():
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    get_anapay_mail(creds)


if __name__ == "__main__":
    main()
