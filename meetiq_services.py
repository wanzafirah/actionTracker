import os
import re
import tempfile

import pandas as pd
import requests
import streamlit as st

try:
    import whisper
except ImportError:
    whisper = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None

from meetiq_constants import OLLAMA_MODEL, WHISPER_MODEL


def get_ollama_url() -> str:
    secret_value = ""
    try:
        secret_value = st.secrets.get("OLLAMA_URL", "")
    except Exception:
        secret_value = ""
    return os.getenv("OLLAMA_URL", secret_value or "http://127.0.0.1:11434/api/generate")


def call_ollama(system: str, user_msg: str, max_tokens: int = 2000) -> str:
    ollama_url = get_ollama_url()
    headers = {}
    if "ngrok" in ollama_url:
        headers["ngrok-skip-browser-warning"] = "true"
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": user_msg,
        "system": system,
        "stream": False,
        "options": {"num_predict": max_tokens, "num_ctx": 2048, "temperature": 0.1, "top_p": 0.9},
    }
    try:
        response = requests.post(ollama_url, json=payload, headers=headers, timeout=300)
        response.raise_for_status()
        return response.json().get("response", "")
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else "unknown"
        detail = exc.response.text[:500] if exc.response is not None else str(exc)
        raise RuntimeError(f"Ollama request failed with status {status_code}. {detail}") from exc
    except requests.RequestException as exc:
        raise RuntimeError(
            f"Could not connect to Ollama at {ollama_url}. If you are deploying on Streamlit Cloud, set OLLAMA_URL to a reachable server."
        ) from exc


@st.cache_resource
def get_local_whisper_model():
    if whisper is None:
        raise RuntimeError("Local Whisper is not installed. Install it with `pip install openai-whisper`.")
    return whisper.load_model(WHISPER_MODEL)


def transcribe_audio_file(uploaded_file, translate_to_english: bool = True) -> str:
    model = get_local_whisper_model()
    file_name = getattr(uploaded_file, "name", "meeting_audio.wav")
    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        raise RuntimeError("Audio file is empty.")
    suffix = os.path.splitext(file_name)[1] or ".wav"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name
    try:
        result = model.transcribe(temp_path, task="translate" if translate_to_english else "transcribe", fp16=False)
        return (result.get("text") or "").strip()
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def dataframe_to_meeting_text(frame: pd.DataFrame, row_limit: int = 40) -> str:
    frame = frame.fillna("")
    if frame.empty:
        return ""
    records = []
    for row_index, row in frame.head(row_limit).iterrows():
        fields = []
        for column in frame.columns:
            value = str(row[column]).strip()
            if value:
                fields.append(f"{column}: {value}")
        if fields:
            records.append(f"Row {row_index + 1}: " + " | ".join(fields))
    return "\n".join(records)


def extract_text_from_document(uploaded_file) -> str:
    file_name = getattr(uploaded_file, "name", "document").lower()
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    if file_name.endswith(".pdf"):
        if PdfReader is None:
            raise RuntimeError("PDF support requires `pypdf`. Install it with `pip install pypdf`.")
        reader = PdfReader(uploaded_file)
        pages = [page.extract_text() or "" for page in reader.pages]
        extracted = "\n".join(page.strip() for page in pages if page.strip())
        if not extracted.strip():
            raise RuntimeError(
                "This PDF has no selectable text. It may be a scanned/image PDF, so the app cannot read it directly."
            )
        return extracted
    if file_name.endswith(".docx"):
        if Document is None:
            raise RuntimeError("Word support requires `python-docx`. Install it with `pip install python-docx`.")
        doc = Document(uploaded_file)
        return "\n".join(paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip())
    if file_name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        chunks = []
        for sheet_name, frame in sheets.items():
            chunks.append(f"Sheet: {sheet_name}")
            chunks.append(dataframe_to_meeting_text(frame))
        return "\n\n".join(chunk for chunk in chunks if chunk.strip())
    if file_name.endswith(".csv"):
        return dataframe_to_meeting_text(pd.read_csv(uploaded_file))
    raise RuntimeError("Unsupported document format. Use PDF, DOCX, XLSX, XLS, or CSV.")
