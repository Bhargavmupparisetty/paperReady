import sys
import atexit
from pathlib import Path
from paperready.config import IDENTITY, ABOUT_TRIGGERS, PAD, W, LEFT, INNER
from paperready.utils import (
    print_section, print_status, print_info, print_ok, print_err,
    _hr, _center_line, _box_line, _box_wrap_lines, ask_permission, open_file_in_app,
    c, C_BLUE, C_GREEN, C_BOLD, C_RESET
)
from paperready.intent import detect_intent, extract_topic, extract_slide_count, detect_summarise_request
from paperready.prompts import build_messages
from paperready.rag import WorkspaceRAG
from paperready.extractors import extract_workspace_file
from paperready.llm import load_model, run_llm_streaming
from paperready.generators import (
    WIN32_OK, create_pptx_via_com, create_docx_via_com,
    create_pptx_fallback, create_docx_fallback, create_txt
)
from paperready.websearch import (
    SELENIUM_OK, DDG_OK, needs_web_search, web_search_text, web_search_images, cleanup_web_images
)

def print_about():
    print()
    print(_hr("="))
    print(_center_line(f"  About {IDENTITY['name']}  "))
    print(_hr("="))
    print(_box_line())
    print(_center_line(f"Designed & Built by  :  {IDENTITY['designer']}"))
    print(_center_line(f"Base Language Model  :  {IDENTITY['base_model']}"))
    print(_center_line(f"Version              :  {IDENTITY['version']}"))
    print(_box_line())
    print(_box_line())
    for wrapped_line in _box_wrap_lines(IDENTITY["purpose"]):
        print(wrapped_line)
    print(_box_line())
    print(_box_line("  What I can do for you:"))
    for cap in IDENTITY["capabilities"]:
        prefix = "    ->  "
        avail = INNER - len(prefix)
        words = cap.split()
        line = ""
        first = True
        for word in words:
            candidate = f"{line} {word}".strip() if line else word
            if len(candidate) > avail:
                if first:
                    print(_box_line(prefix + line))
                    first = False
                else:
                    print(_box_line((" " * len(prefix)) + line))
                line = word
            else:
                line = candidate
        tail = (prefix if first else " " * len(prefix)) + line
        if tail.strip():
            print(_box_line(tail))
    print(_box_line())
    print(_hr("="))
    print()
    if WIN32_OK:
        print_ok("COM automation (pywin32) detected \u2014 PPT/Word will be written LIVE in the app.")
    else:
        print_info("pywin32 not found \u2014 falling back to python-pptx / python-docx file creation.")
        print_info("Install with:  pip install pywin32  then  python -m pywin32_postinstall -install")
    print()

def print_banner():
    from paperready.utils import _figlet_banner
    print()
    print(_hr("="))
    print()
    print(_figlet_banner("PaperReady", font="slant"))
    subtitle = "RAG-Powered Local AI  |  Phi-3 Mini  |  Designed by Bhargav"
    print(PAD + subtitle.center(W))
    print()
    print(_hr("="))
    print(_box_line())
    print(_center_line("COMMANDS"))
    print(_box_line())
    print(_box_line("Type any question                 ->  chat with the AI"))
    print(_box_line("create a 3 slide ppt about X      ->  opens PPT, writes 3 slides"))
    print(_box_line("write a word document about X     ->  opens Word, writes document"))
    print(_box_line("write a text file about X         ->  saves .txt"))
    print(_box_line("summarize report.pptx             ->  reads & summarises workspace file"))
    print(_box_line("summarize notes.docx              ->  reads & summarises workspace file"))
    print(_box_line("who are you / what can you do     ->  about PaperReady"))
    print(_box_line("reload  ->  re-scan workspace/"))
    print(_box_line("exit / quit  ->  close PaperReady"))
    print(_box_line())
    com_status = "COM (live in-app)" if WIN32_OK else "python-pptx / python-docx (file only)"
    if SELENIUM_OK:
        web_status = "Selenium Browser (Chrome/Edge  active)"
    elif DDG_OK:
        web_status = "ddgs fallback (no Selenium)"
    else:
        web_status = "unavailable (pip install selenium)"
    for wl in _box_wrap_lines(f"  PPT/DOCX mode: {com_status}", indent="  "):
        print(wl)
    for wl in _box_wrap_lines(f"  Web search:   {web_status}", indent="  "):
        print(wl)
    print(_box_line())
    print(_hr("="))
    print()

def prompt_user() -> str:
    print(f"{PAD}  {c('[ YOU ]', C_BLUE + C_BOLD)}")
    ans = input(f"{PAD}  {C_GREEN}").strip()
    print(C_RESET, end="", flush=True)
    return ans

def handle_file_output(intent: str, fname: Path, com_app=None, com_doc=None):
    app_names = {"pptx": "Microsoft PowerPoint",
                 "docx": "Microsoft Word",
                 "txt": "Notepad / text editor"}
    app_name = app_names.get(intent, "the default application")
    print_ok(f"File saved  ->  {fname}")
    print_info(f"File size   ->  {fname.stat().st_size // 1024} KB")
    if com_app is not None:
        print_ok(f"{app_name} is open with your content.")
        print_info("The file has been saved. You can continue editing in the app.")
    else:
        if ask_permission(f"Open '{fname.name}' in {app_name}?"):
            if open_file_in_app(fname, app_hint=intent):
                print_ok(f"Opened in {app_name}.")
            else:
                print_err("Auto-open failed. Open manually from outputs/")
        else:
            print_info(f"File is ready at:  {fname}")
            print_info("Open it manually from the outputs/ folder whenever you like.")

def handle_summarise_request(llm, rag: WorkspaceRAG, user_input: str, history: list):
    filename_hint = detect_summarise_request(user_input)
    if not filename_hint:
        print_err("Could not identify a filename in your request.")
        print_info("Try: 'summarize report.pptx' or 'read notes.docx'")
        return history
    target_path = rag.find_file(filename_hint)
    if target_path is None:
        print_err(f"File not found in workspace/:  {filename_hint}")
        print_info("Place the file in the workspace/ folder and type 'reload', then try again.")
        return history
    print_section(f"Reading file  ->  {target_path.name}")
    extracted_text = extract_workspace_file(target_path)
    if not extracted_text.strip():
        print_err(f"Could not extract any text from {target_path.name}.")
        return history
    word_count = len(extracted_text.split())
    print_info(f"Extracted {word_count} words from {target_path.name}")
    MAX_WORDS = 3000
    if word_count > MAX_WORDS:
        truncated = " ".join(extracted_text.split()[:MAX_WORDS])
        print_info(f"(Truncated to {MAX_WORDS} words to fit LLM context window)")
    else:
        truncated = extracted_text
    summarise_query = (
        f"Here is the full content of the file '{target_path.name}':\n\n"
        f"{truncated}\n\n"
        f"Please provide a clear, well-structured summary of this document. "
        f"Identify the main topics, key points, and any important details."
    )
    messages = build_messages(
        history, summarise_query, context="", intent="summarise",
        topic=target_path.stem
    )
    try:
        llm_response = run_llm_streaming(llm, messages)
    except Exception as e:
        print_err(f"Inference error: {e}")
        return history
    history.append({"role": "user", "content": user_input})
    history.append({"role": "assistant", "content": llm_response})
    return history

def main():
    atexit.register(cleanup_web_images)
    print()
    print(_hr("="))
    print(_center_line("PaperReady Editor  --  initialising ..."))
    print(_center_line("Designed by Bhargav  |  Phi-3 Mini (Microsoft)"))
    print(_hr("="))
    print()

    rag = WorkspaceRAG()
    print()

    try:
        llm = load_model()
        print_ok("Model loaded successfully!")
    except Exception as e:
        print_err(f"Model load failed: {e}")
        sys.exit(1)

    print_banner()

    history = []

    while True:
        try:
            user_input = prompt_user()
        except (EOFError, KeyboardInterrupt):
            print(); print_info("Goodbye! \u2014 PaperReady by Bhargav"); break

        if not user_input:
            continue

        if user_input.lower() in ("exit", "quit"):
            print_info("Goodbye! \u2014 PaperReady by Bhargav"); break

        if user_input.lower() == "reload":
            print_section("Reloading Workspace Index"); rag.reload(); continue

        if ABOUT_TRIGGERS.search(user_input):
            print_about()
            history.append({"role": "user", "content": user_input})
            history.append({"role": "assistant", "content": IDENTITY["purpose"]})
            print(_hr("-")); continue

        intent = detect_intent(user_input)
        topic = extract_topic(user_input)
        slide_count = extract_slide_count(user_input) if intent == "pptx" else 6

        if intent == "summarise":
            history = handle_summarise_request(llm, rag, user_input, history)
            print(_hr("-"))
            continue

        if intent != "chat":
            print_section(
                f"Task  ->  {intent.upper()}  |  Topic: {topic}"
                + (f"  |  Slides: {slide_count}" if intent == "pptx" else "")
            )

        context = rag.retrieve_text(user_input)

        web_context = ""
        if needs_web_search(user_input):
            print_section("Web Search")
            web_context = web_search_text(user_input)
            if web_context:
                context = (context + "\n\n--- Live Web Results ---\n" + web_context).strip()

        images = rag.retrieve_images(topic) if intent in ("pptx", "docx") else []

        if images:
            print_info(f"Matched {len(images)} workspace image(s): " + ", ".join(p.name for p in images))
        elif intent in ("pptx", "docx"):
            if DDG_OK:
                print_info("No workspace images found — fetching from web ...")
                images = web_search_images(topic, max_images=3)
            else:
                print_info("No matching images found in workspace/ for this topic.")

        messages = build_messages(history, user_input, context, intent, topic, slide_count)

        try:
            llm_response = run_llm_streaming(llm, messages)
        except Exception as e:
            print_err(f"Inference error: {e}"); continue

        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": llm_response})

        if intent == "pptx":
            print_section(f"Building PowerPoint  ->  {topic}  ({slide_count} slides)")
            preview_lines = [l for l in llm_response.splitlines() if l.strip()][:8]
            print_info("LLM output preview (first 8 non-empty lines):")
            for pl in preview_lines:
                print(f"{PAD}    {pl[:80]}")
            print()
            com_app = com_doc = None
            try:
                if WIN32_OK:
                    fname, com_app, com_doc = create_pptx_via_com(topic, llm_response, images, slide_count)
                else:
                    print_info("pywin32 not available \u2014 using python-pptx fallback.")
                    fname = create_pptx_fallback(topic, llm_response, images, slide_count)
                handle_file_output("pptx", fname, com_app, com_doc)
            except Exception as e:
                print_err(f"PPTX error: {e}")

        elif intent == "docx":
            print_section(f"Building Word Document  ->  {topic}")
            com_app = com_doc = None
            try:
                if WIN32_OK:
                    fname, com_app, com_doc = create_docx_via_com(topic, llm_response, images)
                else:
                    print_info("pywin32 not available \u2014 using python-docx fallback.")
                    fname = create_docx_fallback(topic, llm_response, images)
                handle_file_output("docx", fname, com_app, com_doc)
            except Exception as e:
                print_err(f"DOCX error: {e}")

        elif intent == "txt":
            print_section(f"Writing Text File  ->  {topic}")
            try:
                fname = create_txt(topic, llm_response)
                handle_file_output("txt", fname)
            except Exception as e:
                print_err(f"TXT error: {e}")

        print(_hr("-"))
