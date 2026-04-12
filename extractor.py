"""
Gemini APIを使ってPDFから不動産情報を抽出するモジュール
"""
import os
import json
import logging
from pathlib import Path

from google import genai
from google.genai import types

from config import GEMINI_MODEL, EXTRACTION_PROMPT

logger = logging.getLogger(__name__)

_client: genai.Client | None = None


def configure_gemini(api_key: str | None = None):
    """Gemini APIクライアントを初期化する"""
    global _client
    key = api_key or os.environ.get("GEMINI_API_KEY")
    if not key:
        raise ValueError("GEMINI_API_KEYが設定されていません。.envファイルまたは環境変数を確認してください。")
    _client = genai.Client(api_key=key)


def _get_client() -> genai.Client:
    if _client is None:
        raise RuntimeError("configure_gemini() を先に呼び出してください。")
    return _client


def upload_pdf(pdf_path: str):
    """PDFをGemini Files APIにアップロードする"""
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDFファイルが見つかりません: {pdf_path}")

    client = _get_client()
    logger.info(f"PDFをアップロード中: {path.name}")

    with open(path, "rb") as f:
        uploaded = client.files.upload(
            file=f,
            config={"mime_type": "application/pdf", "display_name": path.name},
        )

    logger.info(f"  アップロード完了: {uploaded.uri}")
    return uploaded


def extract_from_pdfs(pdf_paths: list[str]) -> dict:
    """
    複数のPDFから不動産情報を一括抽出する。
    戻り値: 抽出されたフィールドのdict
    """
    client = _get_client()
    uploaded_files = []

    try:
        for path in pdf_paths:
            uploaded_files.append(upload_pdf(path))

        # アップロードしたファイル＋プロンプトを1リクエストで送信
        contents = uploaded_files + [EXTRACTION_PROMPT]

        logger.info(f"Gemini ({GEMINI_MODEL}) に情報抽出を依頼中...")
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=contents,
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
            ),
        )

        raw = response.text.strip()
        # コードブロックが混入した場合に除去
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]

        extracted = json.loads(raw)
        logger.info(f"  {len(extracted)} フィールドを抽出しました")
        return extracted

    except json.JSONDecodeError as e:
        logger.error(f"JSONパースエラー: {e}")
        logger.error(f"Geminiの応答（先頭500文字）: {response.text[:500]}")
        raise
    finally:
        for f in uploaded_files:
            try:
                client.files.delete(name=f.name)
                logger.debug(f"  一時ファイル削除: {f.name}")
            except Exception:
                pass
