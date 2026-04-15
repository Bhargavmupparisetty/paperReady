import os
import re
from pathlib import Path

ROOT = Path(__file__).parent.parent
WORKSPACE = ROOT / "workspace"
OUTPUTS = ROOT / "outputs"

WORKSPACE.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)
WEB_IMAGES_DIR = OUTPUTS / "web_images"

TEXT_EXTS = {".txt", ".md", ".rst", ".csv", ".log"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
READABLE_DOC_EXTS = {".pptx", ".docx", ".txt", ".md", ".rst", ".csv", ".log"}

W = 86
try:
    _term_width = os.get_terminal_size().columns
except Exception:
    _term_width = 100

if _term_width > W:
    LEFT = (_term_width - W) // 2
else:
    LEFT = 4

PAD = " " * LEFT
INNER = W - 4

IDENTITY = {
    "name": "PaperReady",
    "designer": "Bhargav",
    "base_model": "Phi-3 Mini 4K Instruct by Microsoft",
    "version": "2.0",
    "purpose": (
        "PaperReady is an AI-powered document and presentation assistant "
        "designed and built by Bhargav. It runs a local Phi-3 Mini language model "
        "enhanced with workspace-aware RAG (Retrieval-Augmented Generation). "
        "PaperReady can create rich PowerPoint presentations, formatted Word documents, "
        "and structured text notes \u2014 written to disk and opened directly in PowerPoint "
        "or Word on your PC. It can also embed your own images from the workspace/ folder."
    ),
    "capabilities": [
        "Chat with a local Phi-3 AI \u2014 no internet needed",
        "Create PowerPoint presentations via live COM automation",
        "Create formatted Word (.docx) documents via COM automation",
        "Write plain-text notes (.txt) with auto-header",
        "Insert workspace images into slides and documents automatically",
        "RAG over your workspace/ folder for topic-aware generation",
        "Read and summarise .pptx, .docx, and .txt files from workspace/",
        "Reload workspace index at any time with 'reload'",
        "Tell you about itself when asked",
    ],
}

ABOUT_TRIGGERS = re.compile(
    r"\b(who are you|what are you|about you|tell me about yourself|"
    r"what can you do|your capabilities|what is paperready|"
    r"who (made|built|designed|created) you|your (name|creator|designer|model))\b",
    re.I,
)

SUMMARISE_TRIGGERS = re.compile(
    r"\b(summarize|summarise|summary of|read|open|explain|describe|"
    r"what('?s| is) in|contents? of|give me a summary)\b.{0,60}"
    r"(\.pptx|\.docx|\.txt|\.md|\.csv|\.log|\.rst)\b",
    re.I,
)
