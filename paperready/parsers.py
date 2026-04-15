import re
from datetime import datetime
from paperready.utils import print_info

def _strip_preamble(text: str) -> str:
    trigger = re.compile(
        r"^(SLIDE\s*\d*\s*[:\-–]|#{1,3}\s|\*{2}|\-\s|\*\s|[1-9]\d*[\.\)]|CONTENT\s*[:\-–])",
        re.I | re.M
    )
    m = trigger.search(text)
    if m:
        return text[m.start():]
    return text

def _emergency_slides(topic: str, slide_count: int) -> list:
    t = topic.strip().title() if topic else "This Topic"
    noun = t
    template = [
        (f"{t}: An Overview", [
            f"Introduction to {noun}", f"Why {noun} matters today",
            f"Key principles and scope", f"What this presentation covers"
        ]),
        (f"Understanding {t}", [
            f"Core definition and background of {noun}", f"Historical development and evolution",
            f"Fundamental components explained", f"How {noun} works in practice"
        ]),
        (f"Key Benefits of {t}", [
            f"Primary advantages of {noun}", f"Real-world impact and outcomes",
            f"Efficiency and effectiveness gains", f"Why organisations are adopting {noun}"
        ]),
        (f"Applications of {t}", [
            f"Major industry use cases", f"Practical day-to-day implementations",
            f"Success stories and case studies", f"Emerging opportunities in {noun}"
        ]),
        (f"Challenges & Considerations", [
            f"Common obstacles when applying {noun}", f"Risk factors to manage",
            f"Implementation and adoption hurdles", f"Strategies to overcome limitations"
        ]),
        (f"The Future of {t}", [
            f"Emerging trends shaping {noun}", f"Predictions for the next five years",
            f"New opportunities on the horizon", f"How to prepare for what's next"
        ]),
        (f"Key Takeaways", [
            f"{noun} is transforming the landscape", f"Summary of concepts covered today",
            f"Recommended next steps and actions", f"Thank you \u2014 questions welcome"
        ]),
    ]
    result = template[:slide_count]
    idx = len(result) + 1
    while len(result) < slide_count:
        result.append((f"{t} \u2014 Insights Part {idx}", [
            f"Additional analysis of {noun}", f"Deeper dive into key concepts",
            f"Further considerations and context", f"Continued exploration of {noun}"
        ]))
        idx += 1
    return result[:slide_count]

def parse_slides(llm_text: str, topic: str, slide_count: int) -> list:
    clean_text = _strip_preamble(llm_text)
    lines = clean_text.splitlines()
    slides_data = []
    cur_title = ""
    cur_bullets = []
    found_strict = False
    for line in lines:
        line = line.strip()
        m_slide = re.match(r"SLIDE\s*\d*\s*[:\-–]\s*(.+)", line, re.I)
        m_cont = re.match(r"CONTENT\s*[:\-–]\s*(.+)", line, re.I)
        if m_slide:
            if cur_title:
                slides_data.append((cur_title, cur_bullets))
            cur_title = m_slide.group(1).strip().strip("*_")
            cur_bullets = []
            found_strict = True
        elif m_cont and found_strict:
            cur_bullets = [b.strip() for b in m_cont.group(1).split("|") if b.strip()]
    if cur_title:
        slides_data.append((cur_title, cur_bullets))
    if slides_data and _has_real_content(slides_data):
        return _pad_slides(slides_data, topic, slide_count)
    
    slides_data = []
    cur_title = ""
    cur_bullets = []
    
    def _is_heading(l: str) -> str:
                                                     
        m = re.match(r"^[Ss]lide\s*\d*\s*[:\-–]?\s*(.+)", l)
        if m: return m.group(1).strip("*_ ")
                          
        m = re.match(r"^#{1,4}\s+(.+)", l)
        if m: return m.group(1).strip("*_ ")
                           
        m = re.match(r"^\*{2}(.+?)\*{2}\s*$", l)
        if m: return m.group(1).strip()
                                     
        m = re.match(r"^([A-Z][a-zA-Z\s]+):$", l)
        if m: return m.group(1).strip()
        return ""

    bullet_re = re.compile(r"^[\-\*\u2022\u25b8•]\s+(.+)")
    numbered_re = re.compile(r"^\d+[\.\)]\s+(.+)")
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        ht = _is_heading(line)
        if ht and "|" not in line and len(ht) < 80:
            if cur_title or cur_bullets:
                slides_data.append((cur_title or topic, cur_bullets))
            cur_title = ht
            cur_bullets = []
        else:
            mb = bullet_re.match(line) or numbered_re.match(line)
            if mb:
                if not cur_title:
                    cur_title = topic
                cur_bullets.append(mb.group(1).strip())
            elif "|" in line and cur_title:
                cur_bullets.extend([b.strip() for b in line.split("|") if b.strip()])
            else:
                                                                             
                if cur_title and len(line) > 10:
                    cur_bullets.append(line.strip("*_ "))
                    
    if cur_title or cur_bullets:
        slides_data.append((cur_title or topic, cur_bullets))
        
    if slides_data and _has_real_content(slides_data):
        return _pad_slides(slides_data, topic, slide_count)
    
    paragraphs = [p.strip() for p in re.split(r"\n{2,}", clean_text) if p.strip()]
    slides_data = []
    for i, para in enumerate(paragraphs[:slide_count]):
        sents = re.split(r"(?<=[.!?])\s+", para)
        title = sents[0][:60] if sents else f"Slide {i + 1}"
        buls = [s[:120] for s in sents[1:5]] if len(sents) > 1 else [para[:120]]
        slides_data.append((title, buls))
    if slides_data and _has_real_content(slides_data):
        return _pad_slides(slides_data, topic, slide_count)
    print_info("LLM output did not match any format \u2014 using topic-aware emergency content.")
    return _emergency_slides(topic, slide_count)

def _has_real_content(slides: list) -> bool:
    good = 0
    for _title, bullets in slides:
        if any(len(b.strip()) > 15 for b in bullets):
            good += 1
    return good >= max(1, len(slides) // 2)

def _pad_slides(slides: list, topic: str, target: int) -> list:
    if slides and not slides[0][1]:
        slides[0] = (slides[0][0], [topic, datetime.now().strftime("%B %Y")])
    _filler = re.compile(r"^\s*(sure|happy|glad|here|let me|i'll|of course|certainly)[^a-z]*$", re.I)
    if slides and _filler.match(slides[0][0]):
        slides[0] = (topic.title(), slides[0][1])
    if len(slides) < target:
        emergency = _emergency_slides(topic, target)
        while len(slides) < target:
            slides.append(emergency[len(slides)])
    return slides[:target]
