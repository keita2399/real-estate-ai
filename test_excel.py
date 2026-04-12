"""
Gemini APIなしでExcel書き込みのみをテストするスクリプト
"""
from create_template import create_template
from excel_writer import write_to_excel, preview_extracted

# 法務省サンプルを模したモックデータ
MOCK_DATA = {
    "所在":             "東京都新宿区西新宿二丁目",
    "地番":             "1番1",
    "地目":             "宅地",
    "地積":             "245.50",
    "建物所在":         "東京都新宿区西新宿二丁目1番地1",
    "家屋番号":         "1番1の101",
    "種類":             "居宅",
    "構造":             "鉄筋コンクリート造陸屋根3階建",
    "床面積":           "82.34",
    "建築年月":         "平成15年3月",
    "所有者氏名":       "山田太郎",
    "所有者住所":       "東京都杉並区荻窪一丁目2番3号",
    "取得原因":         "売買",
    "取得日":           "令和3年5月20日",
    "共有持分":         "記載なし",
    "抵当権有無":       "あり",
    "債権額":           "3,500万円",
    "債権者":           "株式会社〇〇銀行",
    "債務者":           "山田太郎",
    "設定日":           "令和3年5月20日",
    "用途地域":         "第二種住居地域",
    "建ぺい率":         "60",
    "容積率":           "300",
    "防火地域":         "準防火地域",
    "高度地区":         "第三種高度地区",
    "上水道":           "公営水道（引込済み）",
    "下水道":           "公共下水道（接続済み）",
    "ガス":             "都市ガス（東京ガス）",
    "電気":             "東京電力エナジーパートナー",
    "洪水浸水想定区域": "該当なし",
    "土砂災害警戒区域": "該当なし",
    "津波災害警戒区域": "該当なし",
    "液状化リスク":     "低",
    "避難場所":         "新宿中央公園",
}

if __name__ == "__main__":
    # テンプレート確認（なければ生成）
    import os
    if not os.path.exists("template/jyuujiku.xlsx"):
        create_template("template/jyuujiku.xlsx")

    preview_extracted(MOCK_DATA)

    output_path = write_to_excel(
        MOCK_DATA,
        template_path="template/jyuujiku.xlsx",
        output_dir="output",
        prefix="テストデータ_新宿区",
    )
    print(f"出力: {output_path}")
