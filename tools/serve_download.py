#!/usr/bin/env python3
"""خادم تحميل صغير: صفحة تحميل + الملفات (PDF/DOCX).

يُشغَّل من جذر المستودع:  python3 tools/serve_download.py [port]
يخدم مجلد output/site على 0.0.0.0 ليكون متاحًا للمعاينة/التحميل من الخارج.
"""
import http.server
import os
import socketserver
import sys

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "site")
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8080


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def end_headers(self):
        # لا قيود على التضمين، ونوع محتوى صريح للتنزيل
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    def log_message(self, fmt, *args):  # هدوء السجلّ
        sys.stderr.write("· %s\n" % (fmt % args))


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True


if __name__ == "__main__":
    os.chdir(ROOT)
    with Server(("0.0.0.0", PORT), Handler) as httpd:
        print(f"صفحة التحميل على المنفذ {PORT} — المجلد: {ROOT}", flush=True)
        httpd.serve_forever()
