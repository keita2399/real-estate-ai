"""
不動産書類AI自動入力システム
使用方法:
    python main.py input/touki_sample.pdf
    python main.py input/        # フォルダ内の全PDFを処理
    python main.py input/touki.pdf input/hazard.pdf   # 複数ファイル
"""
import os
import sys
import logging
import argparse
from pathlib import Path
from dotenv import load_dotenv

from extractor import configure_gemini, extract_from_pdfs
from excel_writer import write_to_excel, preview_extracted
from create_template import create_template

# ---- ロギング設定 ----
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

TEMPLATE_PATH = "template/jyuujiku.xlsx"


def collect_pdfs(inputs: list[str]) -> list[str]:
    """引数からPDFパスのリストを収集する"""
    pdfs = []
    for inp in inputs:
        p = Path(inp)
        if p.is_dir():
            pdfs.extend(sorted(p.glob("*.pdf")))
        elif p.suffix.lower() == ".pdf" and p.exists():
            pdfs.append(p)
        else:
            logger.warning(f"スキップ（PDFではないかファイルなし）: {inp}")
    return [str(p) for p in pdfs]


def ensure_template():
    """テンプレートが存在しない場合は自動生成する"""
    if not Path(TEMPLATE_PATH).exists():
        logger.info("Excelテンプレートが見つかりません。自動生成します...")
        create_template(TEMPLATE_PATH)


def run(pdf_paths: list[str], output_dir: str = "output", prefix: str = ""):
    """メイン処理"""
    if not pdf_paths:
        logger.error("処理対象のPDFが見つかりません。")
        sys.exit(1)

    logger.info(f"処理開始: {len(pdf_paths)} ファイル")
    for p in pdf_paths:
        logger.info(f"  - {p}")

    # Gemini設定
    load_dotenv()
    configure_gemini()

    # テンプレート確認
    ensure_template()

    # PDF → Gemini抽出
    extracted = extract_from_pdfs(pdf_paths)

    # ターミナルプレビュー
    preview_extracted(extracted)

    # Excel書き込み
    output_path = write_to_excel(
        extracted,
        template_path=TEMPLATE_PATH,
        output_dir=output_dir,
        prefix=prefix or Path(pdf_paths[0]).stem,
    )

    print(f"\n[完了] ファイルを保存しました:\n  {output_path}\n")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="不動産書類PDF → 重要事項説明書Excel 自動入力ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python main.py input/touki_sample.pdf
  python main.py input/
  python main.py input/touki.pdf input/hazard.pdf -o output/ -p 物件001
        """,
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        help="PDFファイルまたはPDFが入ったフォルダ（複数指定可）",
    )
    parser.add_argument(
        "-o", "--output",
        default="output",
        help="出力先ディレクトリ（デフォルト: output/）",
    )
    parser.add_argument(
        "-p", "--prefix",
        default="",
        help="出力ファイル名のプレフィックス",
    )
    parser.add_argument(
        "--create-template-only",
        action="store_true",
        help="Excelテンプレートだけを生成して終了",
    )

    args = parser.parse_args()

    if args.create_template_only:
        create_template(TEMPLATE_PATH)
        return

    pdfs = collect_pdfs(args.inputs)
    run(pdfs, output_dir=args.output, prefix=args.prefix)


if __name__ == "__main__":
    main()
