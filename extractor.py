"""
LangChain + Gemini APIを使ってPDFから不動産情報を抽出するモジュール
"""
import os
import json
import logging
from pathlib import Path

from google import genai
from google.genai import types
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage

from config import GEMINI_MODEL, EXTRACTION_PROMPT

logger = logging.getLogger(__name__)

_gemini_client: genai.Client | None = None
_langchain_model: ChatGoogleGenerativeAI | None = None


def configure_gemini(api_key: str | None = None):
    """Gemini APIクライアントを初期化する"""
    global _gemini_client, _langchain_model
    key = api_key or os.environ.get("GEMINI_API_KEY")
    if not key:
        raise ValueError("GEMINI_API_KEYが設定されていません。.envファイルまたは環境変数を確認してください。")
    _gemini_client = genai.Client(api_key=key)
    _langchain_model = ChatGoogleGenerativeAI(
        model=GEMINI_MODEL,
        google_api_key=key,
        temperature=0.0,
    )


def _get_client() -> genai.Client:
    if _gemini_client is None:
        raise RuntimeError("configure_gemini() を先に呼び出してください。")
    return _gemini_client


def _get_model() -> ChatGoogleGenerativeAI:
    if _langchain_model is None:
        raise RuntimeError("configure_gemini() を先に呼び出してください。")
    return _langchain_model


def upload_pdf(pdf_path: str):
    """PDFをGemini Files APIにアップロードする（ファイルAPIはgoogle-genaiを継続使用）"""
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
    複数のPDFから不動産情報を一括抽出する（LangChain経由）。
    戻り値: 抽出されたフィールドのdict
    """
    model = _get_model()
    client = _get_client()
    uploaded_files = []

    try:
        for path in pdf_paths:
            uploaded_files.append(upload_pdf(path))

        # LangChain の HumanMessage でファイルURIとプロンプトを送信
        file_content = []
        for uploaded in uploaded_files:
            file_content.append({
                "type": "media",
                "file_uri": uploaded.uri,
                "mime_type": "application/pdf",
            })
        file_content.append({
            "type": "text",
            "text": EXTRACTION_PROMPT,
        })

        logger.info(f"LangChain + Gemini ({GEMINI_MODEL}) に情報抽出を依頼中...")
        response = model.invoke([HumanMessage(content=file_content)])

        raw = (response.content if isinstance(response.content, str) else str(response.content)).strip()

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
        raise
    finally:
        for f in uploaded_files:
            try:
                client.files.delete(name=f.name)
                logger.debug(f"  一時ファイル削除: {f.name}")
            except Exception:
                pass
