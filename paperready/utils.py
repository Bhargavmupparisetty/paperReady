import os
import platform
import subprocess
from pathlib import Path

try:
    import pyfiglet
    FIGLET_OK = True
except ImportError:
    FIGLET_OK = False

from paperready.config import PAD, W, LEFT, INNER

# --- ANSI Colors ---
C_RESET = "\033[0m"
C_BOLD = "\033[1m"
C_CYAN = "\033[96m"
C_GREEN = "\033[92m"
C_YELLOW = "\033[93m"
C_RED = "\033[91m"
C_MAG = "\033[95m"
C_BLUE = "\033[94m"
C_DIM = "\033[2m"

def c(text: str, color: str) -> str:
    return f"{color}{text}{C_RESET}"


def _hr(char: str = "=") -> str:
    return PAD + char * W

def _box_line(text: str = "", char: str = "|") -> str:
    if len(text) > INNER:
        text = text[:INNER - 3] + "..."
    return PAD + char + " " + text.ljust(INNER) + " " + char

def _box_wrap_lines(text: str, indent: str = "", char: str = "|") -> list:
    available = INNER - len(indent)
    words = text.split()
    rows = []
    line = ""
    for word in words:
        candidate = f"{line} {word}".strip() if line else word
        if len(candidate) > available:
            if line:
                rows.append(_box_line(indent + line, char))
            line = word
        else:
            line = candidate
    if line:
        rows.append(_box_line(indent + line, char))
    return rows if rows else [_box_line(indent, char)]

def _center_line(text: str, char: str = "|") -> str:
    return PAD + char + " " + text.center(INNER) + " " + char

def _label(tag: str, text: str) -> str:
    tag_str = f"[{tag:^7}]"
    return f"{PAD}{tag_str}  {text}"

def print_section(title: str):
    print()
    print(c(_hr("-"), C_CYAN))
    print(_label(c("  >>  ", C_BOLD + C_CYAN), c(title, C_BOLD)))
    print(c(_hr("-"), C_CYAN))

def print_status(symbol: str, message: str):
    print(f"{PAD}  {symbol}  {message}")

def print_info(message: str):
    print(f"{PAD}  {c('*', C_YELLOW)}  {message}")

def print_ok(message: str):
    print(f"{PAD}  {c('[OK]', C_GREEN)}  {message}")

def print_err(message: str):
    print(f"{PAD}  {c('[!!]', C_RED)}  {message}")

def _figlet_banner(text: str, font: str = "slant") -> str:
    if FIGLET_OK:
        try:
            raw = pyfiglet.figlet_format(text, font=font)
            lines = raw.splitlines()
            return "\n".join(PAD + c(line, C_CYAN + C_BOLD) for line in lines)
        except Exception:
            pass
    border = PAD + c("+" + "-" * (W - 2) + "+", C_CYAN)
    inner = W - 4
    middle = PAD + c("|", C_CYAN) + c(text.center(inner + 2), C_BOLD) + c("|", C_CYAN)
    return "\n".join([border, middle, border])

def ask_permission(action: str) -> bool:
    print()
    print(f"{PAD}  [PERMISSION]  {action}")
    print(f"{PAD}  Allow? (yes / no)  ->  ", end="", flush=True)
    try:
        ans = input().strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return ans in ("yes", "y", "1", "ok", "sure", "yep", "yeah")

def open_file_in_app(file_path: Path, app_hint: str = "auto") -> bool:
    system = platform.system()
    path_str = str(file_path.resolve())
    try:
        if system == "Windows":
            if app_hint == "txt":
                subprocess.Popen(["notepad.exe", path_str])
            else:
                os.startfile(path_str)
        elif system == "Darwin":
            subprocess.Popen(["open", path_str])
        else:
            subprocess.Popen(["xdg-open", path_str])
        return True
    except Exception as e:
        print_err(f"Could not open file automatically: {e}")
        print_info(f"Please open manually:  {path_str}")
        return False
