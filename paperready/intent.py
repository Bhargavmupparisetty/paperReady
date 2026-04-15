import re

_INTENT_PATTERNS = {
    "pptx": re.compile(r"\b(powerpoint|pptx|presentation|slides?|slide deck)\b", re.I),
    "docx": re.compile(r"\b(word\s+doc(ument)?|docx|\.docx)\b", re.I),
    "txt": re.compile(r"\b(notepad|text\s*file|\.txt|write\s+(to\s+)?file|save\s+(as\s+)?text)\b", re.I),
    "summarise": re.compile(
        r"\b(summarize|summarise|summary|read|open|explain|describe|"
        r"what('?s| is) in|contents? of|give me a summary)\b",
        re.I,
    ),
}

def detect_summarise_request(query: str) -> str | None:
    m = re.search(r"([\w\-\.]+\.(pptx|docx|txt|md|csv|log|rst))", query, re.I)
    return m.group(1) if m else None

def detect_intent(query: str) -> str:
    if _INTENT_PATTERNS["summarise"].search(query) and detect_summarise_request(query):
        return "summarise"
    for intent, pattern in _INTENT_PATTERNS.items():
        if intent == "summarise":
            continue
        if pattern.search(query):
            return intent
    return "chat"

def extract_topic(query: str) -> str:
    q = query.strip()
                                                          
    q = re.sub(
        r",\s*(?:and\s+)?(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|"
        r"[1-9]\d*(?:st|nd|rd|th)?)\s+slide\b[^,]*",
        "", q, flags=re.I
    )
                                
    q = re.sub(r"\bin\s+\d+[\s-]*slides?\b", " ", q, flags=re.I)
    q = re.sub(r"\b\d+[\s-]*slides?\b", " ", q, flags=re.I)
    for kw in ["powerpoint", "pptx", "ppt", "presentation", "slides", "slide deck",
               "word document", "word doc", "docx", "notepad", "text file",
               "write a", "create a", "make a", "generate a",
               "write", "create", "make", "generate",
               "about", "on", "for", "regarding", "please", "me", "a"]:
        q = re.sub(r"\b" + re.escape(kw) + r"\b", " ", q, flags=re.I)
    topic = re.sub(r"\s+", " ", q).strip().strip(".,;:'\"")
    topic = re.sub(
        r"\s+\b(in|on|for|the|an|a|of|and|or|with|by|to|into|from|about|at|as)\b\s*$",
        "", topic, flags=re.I
    ).strip().strip(".,;:'\"")
    return topic if topic else "My Topic"

def extract_slide_count(query: str) -> int:
    word_map = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    }
    m = re.search(r"\b(\d+)[\s-]*slides?\b", query, re.I)
    if m:
        return max(2, min(20, int(m.group(1))))
    for word, num in word_map.items():
        if re.search(r"\b" + word + r"[\s-]*slides?\b", query, re.I):
            return num
    return 6
