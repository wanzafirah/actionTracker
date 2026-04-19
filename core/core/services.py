import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from meetiq_services import call_ollama, extract_text_from_document, get_ollama_url, normalize_ollama_url, transcribe_audio_file

__all__ = [
    "call_ollama",
    "extract_text_from_document",
    "get_ollama_url",
    "normalize_ollama_url",
    "transcribe_audio_file",
]
