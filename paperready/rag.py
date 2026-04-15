import re
import math
from pathlib import Path
from collections import defaultdict
from paperready.config import WORKSPACE, TEXT_EXTS, IMAGE_EXTS, READABLE_DOC_EXTS
from paperready.utils import print_info
from paperready.extractors import extract_workspace_file

class WorkspaceRAG:
    def __init__(self):
        self.text_docs: dict = {}
        self.image_paths: list = []
        self.tfidf: dict = {}
        self._build_index()

    def _build_index(self):
        for p in WORKSPACE.rglob("*"):
            if p.is_file():
                if p.suffix.lower() in TEXT_EXTS:
                    try:
                        self.text_docs[p.name] = p.read_text(encoding="utf-8", errors="ignore")
                    except Exception:
                        pass
                elif p.suffix.lower() in IMAGE_EXTS:
                    self.image_paths.append(p)
                elif p.suffix.lower() in {".pptx", ".docx"}:
                    extracted = extract_workspace_file(p)
                    if extracted:
                        self.text_docs[p.name] = extracted
        if self.text_docs:
            self._compute_tfidf()
        print_info(
            f"RAG index built  ->  {len(self.text_docs)} text/doc file(s), "
            f"{len(self.image_paths)} image(s) in workspace/"
        )

    def _tokenize(self, text: str) -> list:
        return re.findall(r"[a-z]+", text.lower())

    def _compute_tfidf(self):
        tf = defaultdict(lambda: defaultdict(int))
        df = defaultdict(int)
        N = len(self.text_docs)
        for name, content in self.text_docs.items():
            words = self._tokenize(content)
            for w in words:
                tf[name][w] += 1
            for w in set(words):
                df[w] += 1
        for name, word_counts in tf.items():
            total = sum(word_counts.values()) or 1
            for w, count in word_counts.items():
                idf = math.log((N + 1) / (df[w] + 1)) + 1
                if w not in self.tfidf:
                    self.tfidf[w] = {}
                self.tfidf[w][name] = (count / total) * idf

    def retrieve_text(self, query: str, top_k: int = 3, max_chars: int = 1200) -> str:
        if not self.text_docs:
            return ""
        q_words = set(self._tokenize(query))
        scores = defaultdict(float)
        for w in q_words:
            if w in self.tfidf:
                for docname, score in self.tfidf[w].items():
                    scores[docname] += score
        ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        snippets = []
        for docname, _ in ranked[:top_k]:
            text = self.text_docs[docname]
            snippets.append(f"[From: {docname}]\n{text[:max_chars]}")
        return "\n\n".join(snippets)

    def retrieve_images(self, query: str) -> list:
        q_words = set(self._tokenize(query))
        q_phrase = re.sub(r"\s+", "_", query.strip().lower())
        matched = []
        for img in self.image_paths:
            stem_words = set(self._tokenize(img.stem))
            if q_words & stem_words or q_phrase in img.stem.lower():
                matched.append(img)
        return matched

    def reload(self):
        self.text_docs.clear()
        self.image_paths.clear()
        self.tfidf.clear()
        self._build_index()

    def find_file(self, filename_hint: str) -> Path | None:
        hint_lower = filename_hint.strip().lower()
        for p in WORKSPACE.rglob("*"):
            if p.is_file() and p.suffix.lower() in READABLE_DOC_EXTS:
                if hint_lower in p.name.lower():
                    return p
        return None
