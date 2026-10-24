"""
特別価格回答データの列定義（テンプレート作成と集約で共通に使う）
列名を変えるとテンプレートと集約の両方に反映される
"""

TABLE_NAME = "tbl_price"
DATA_SHEET = "データ"

# (列名, 型) 型は str / date / num
COLUMNS = [
    ("案件ID", "str"),
    ("案件名", "str"),
    ("顧客名", "str"),
    ("回答日", "date"),
    ("有効期限", "date"),
    ("品番", "str"),
    ("品名", "str"),
    ("数量", "num"),
    ("単位", "str"),
    ("標準価格", "num"),
    ("特別価格", "num"),
    ("値引率", "num"),
    ("通貨", "str"),
    ("担当者", "str"),
    ("承認者", "str"),
    ("理由・備考", "str"),
]

COLUMN_NAMES = [name for name, _ in COLUMNS]
