import re
import time
import shutil
import requests
from pathlib import Path
from paperready.config import WEB_IMAGES_DIR
from paperready.utils import print_info, print_err, print_ok

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

try:
    from ddgs import DDGS
    DDG_OK = True
except ImportError:
    DDG_OK = False

_CURRENT_KEYWORDS = re.compile(
    r"\b(latest|current|recent|today|now|news|2024|2025|2026|"
    r"trending|just|breaking|live|update|this week|this month|"
    r"right now|happening|announced|released|launched|won|lost|"
    r"score|match|election|war|crisis|event|ipl|nba|fifa|"
    r"championship|tournament|season|market|stock|price)\b",
    re.I,
)

def needs_web_search(query: str) -> bool:
    return bool(_CURRENT_KEYWORDS.search(query))

def _open_browser():
    if not SELENIUM_OK:
        return None
                      
    try:
        opts = ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = webdriver.Chrome(options=opts)
        return driver
    except Exception:
        pass
                      
    try:
        opts = EdgeOptions()
        opts.add_argument("--start-maximized")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = webdriver.Edge(options=opts)
        return driver
    except Exception:
        return None

def web_search_text(query: str, max_results: int = 5) -> str:
    driver = _open_browser()
    if driver:
        try:
            print_info(f"Browser search:  {query[:70]}")
            search_url = f"https://www.bing.com/search?q={requests.utils.quote(query)}"
            driver.get(search_url)
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.b_algo"))
            )
            results = driver.find_elements(By.CSS_SELECTOR, "li.b_algo")
            snippets = []
            for r in results[:max_results]:
                try:
                    title = r.find_element(By.CSS_SELECTOR, "h2").text.strip()
                    try:
                        body = r.find_element(By.CSS_SELECTOR, "p").text.strip()
                    except Exception:
                        body = ""
                    try:
                        href = r.find_element(By.CSS_SELECTOR, "h2 a").get_attribute("href") or ""
                    except Exception:
                        href = ""
                    if title:
                        snippets.append(f"[{title}]\n{body}\nSource: {href}")
                except Exception:
                    continue
            driver.quit()
            if snippets:
                print_ok(f"Browser returned {len(snippets)} result(s).")
                return "\n\n".join(snippets)
        except Exception as e:
            print_err(f"Browser text search failed: {e}")
            try:
                driver.quit()
            except Exception:
                pass

                   
    if DDG_OK:
        try:
            print_info("Falling back to ddgs ...")
            results = []
            with DDGS() as ddgs:
                for r in ddgs.text(query, max_results=max_results):
                    title = r.get("title", "")
                    body  = r.get("body", "")
                    href  = r.get("href", "")
                    if title or body:
                        results.append(f"[{title}]\n{body}\nSource: {href}")
            if results:
                print_ok(f"ddgs returned {len(results)} result(s).")
                return "\n\n".join(results)
        except Exception as e:
            print_err(f"ddgs fallback failed: {e}")
    return ""

def web_search_images(query: str, max_images: int = 3) -> list:
    WEB_IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    downloaded = []
    image_urls = []

    driver = _open_browser()
    if driver:
        try:
            print_info(f"Browser image search:  {query[:60]}")
            search_url = f"https://www.bing.com/images/search?q={requests.utils.quote(query)}"
            driver.get(search_url)
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "img.mimg"))
            )
                                            
            driver.execute_script("window.scrollTo(0, 600);")
            time.sleep(1.5)
            imgs = driver.find_elements(By.CSS_SELECTOR, "img.mimg")
            for img in imgs:
                src = img.get_attribute("src") or img.get_attribute("data-src") or ""
                if src.startswith("http") and src not in image_urls:
                    image_urls.append(src)
                if len(image_urls) >= max_images * 2:
                    break
            driver.quit()
            print_ok(f"Browser found {len(image_urls)} candidate image(s).")
        except Exception as e:
            print_err(f"Browser image search failed: {e}")
            try:
                driver.quit()
            except Exception:
                pass

                                                 
    if not image_urls and DDG_OK:
        try:
            print_info("Falling back to ddgs images ...")
            with DDGS() as ddgs:
                results = list(ddgs.images(query, max_results=max_images * 3))
            for r in results:
                url = r.get("image", "")
                if url:
                    image_urls.append(url)
        except Exception as e:
            print_err(f"ddgs image fallback failed: {e}")

                             
    headers = {"User-Agent": "Mozilla/5.0"}
    count = 0
    for url in image_urls:
        if count >= max_images:
            break
        ext = Path(url.split("?")[0]).suffix.lower()
        if ext not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
            ext = ".jpg"
        safe_name = re.sub(r"[^\w]", "_", query[:30])
        dest = WEB_IMAGES_DIR / f"{safe_name}_{count}{ext}"
        try:
            resp = requests.get(url, stream=True, timeout=8, headers=headers)
            resp.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            downloaded.append(dest)
            print_ok(f"  Image saved: {dest.name}")
            count += 1
        except Exception:
            continue

    return downloaded

def cleanup_web_images():
    if WEB_IMAGES_DIR.exists():
        shutil.rmtree(WEB_IMAGES_DIR, ignore_errors=True)
