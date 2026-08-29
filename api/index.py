from http.server import BaseHTTPRequestHandler
import subprocess

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        # Note: Streamlit n'est pas nativement supporté par le runtime serverless de Vercel.
        # Pour une application Streamlit, il est fortement recommandé d'utiliser 
        # "Streamlit Community Cloud" (gratuit) au lieu de Vercel.
        self.wfile.write(b"Streamlit apps are best hosted on Streamlit Community Cloud.")
        return