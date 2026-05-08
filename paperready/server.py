import http.server
import socketserver
import threading
import json
import webbrowser
from pathlib import Path

PORT = 8080
UI_STATE = {
    "topic": "Welcome to PaperReady Canvas",
    "html": "<p>Waiting for commands...</p>",
    "graphviz": ""
}

class EditorHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            ui_path = Path(__file__).parent / "editor_ui.html"
            with open(ui_path, "rb") as f:
                self.wfile.write(f.read())
        elif self.path == '/state':
            self.send_response(200)
            self.send_header("Content-type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(UI_STATE).encode('utf-8'))
        else:
            self.send_error(404)

    def log_message(self, format, *args):
        pass

def start_server():
    handler = EditorHandler
    global PORT
    for p in range(8080, 8100):
        try:
            httpd = socketserver.TCPServer(("", p), handler)
            PORT = p
            break
        except OSError:
            continue
    else:
        return None

    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    return PORT

def update_ui_state(topic: str, html_content: str, graphviz_code: str = ""):
    global UI_STATE
    UI_STATE["topic"] = topic
    UI_STATE["html"] = html_content
    UI_STATE["graphviz"] = graphviz_code

def open_ui():
    webbrowser.open(f"http://localhost:{PORT}/")
