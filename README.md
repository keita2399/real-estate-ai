# 不動産書類AI自動入力システム

不動産書類（PDF）をGemini AIで解析し、重要事項説明書ExcelテンプレートへAIが自動入力するツールです。

## 概要

```
PDF（登記簿謄本・ハザードマップ等）
        ↓
   Gemini AI（情報抽出）
        ↓
重要事項説明書 Excel（自動入力）
        ↓
  宅建士による確認・修正（約20分）
```

## セットアップ

```bash
pip install -r requirements.txt
cp .env.example .env
# .env に GEMINI_API_KEY を設定
```

## 使い方

```bash
# 1ファイル
python main.py input/touki_sample.pdf

# フォルダ内の全PDF
python main.py input/

# 複数ファイルを1物件として処理
python main.py input/touki.pdf input/hazard.pdf input/tosui.pdf

# 出力ファイル名プレフィックス指定
python main.py input/ -p 物件001_渋谷区

# テンプレートだけ再生成
python main.py --create-template-only
```

## 出力

`output/` フォルダに `重要事項説明書_YYYYMMDD_HHMMSS.xlsx` が生成されます。

## 対応書類

| 書類 | 抽出できる情報 |
|------|--------------|
| 登記事項証明書（登記簿謄本） | 所在・地番・地目・地積・所有者・抵当権 |
| ハザードマップ | 洪水・土砂災害・津波リスク |
| 都市計画図 | 用途地域・建ぺい率・容積率 |
| 上下水道図 | 上水道・下水道の状況 |

## 自動入力される項目（計33項目）

**物件基本情報**（土地・建物）、**所有権情報**、**抵当権情報**、
**都市計画法制限**、**インフラ状況**、**ハザードマップ情報**

## 注意事項

- AI抽出結果は必ず宅地建物取引士が確認・修正してください
- PDFのスキャン品質が低い場合、精度が落ちることがあります
- 入力内容はサーバーに送信されません（Gemini Files APIへのアップロード後、自動削除）
