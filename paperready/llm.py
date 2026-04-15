import psutil
from paperready.utils import print_info, LEFT, PAD, W, c, C_BOLD, C_RESET, C_MAG

try:
    from llama_cpp import Llama
    LLAMA_OK = True
except ImportError:
    LLAMA_OK = False

def load_model():
    if not LLAMA_OK:
        raise RuntimeError("llama_cpp not installed.  Run:  pip install llama-cpp-python")
    print_info("Loading Phi-3 Mini Q4 model (~2.2 GB) from Hugging Face cache ...")
    print_info("First run will download; subsequent runs use local cache.")
    llm = Llama.from_pretrained(
        repo_id="microsoft/Phi-3-mini-4k-instruct-gguf",
        filename="Phi-3-mini-4k-instruct-q4.gguf",
        n_ctx=4096,
        n_threads=min(8, psutil.cpu_count(logical=False) or 4),
        verbose=False,
    )
    return llm

def run_llm_streaming(llm, messages: list) -> str:
    print()
    print(f"{PAD}  {c('[ AI ]', C_MAG + C_BOLD)}")
    print(f"{PAD}  {C_RESET}", end="", flush=True)

    chunks = llm.create_chat_completion(
        messages=messages,
        stream=True,
        max_tokens=2048,
        temperature=0.4,
    )
    full = ""
    col = LEFT + 2
    wrap_at = LEFT + W - 2

    for chunk in chunks:
        delta = chunk["choices"][0]["delta"]
        if "content" in delta:
            piece = delta["content"]
            for char in piece:
                if char == "\n":
                    print(); print(f"{PAD}  ", end="", flush=True); col = LEFT + 2
                else:
                    print(char, end="", flush=True); col += 1
                    if col >= wrap_at:
                        print(); print(f"{PAD}  ", end="", flush=True); col = LEFT + 2
            full += piece

    print(C_RESET); print()
    return full
