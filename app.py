from __future__ import annotations

import io
import json
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PurePosixPath
from urllib.parse import urlparse

from converter import convert

ROOT = Path(__file__).parent
MAX_UPLOAD = 40 * 1024 * 1024
MAX_UNPACKED = 100 * 1024 * 1024
MAX_FILES = 500
IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.pdf'}
PROJECT_EXTENSIONS = IMAGE_EXTENSIONS | {'.tex', '.bib'}


def read_upload(headers, body):
    content_type = headers.get('Content-Type', '')
    if not content_type.startswith('multipart/form-data'):
        raise ValueError('Upload a LaTeX project using the form.')
    message = BytesParser(policy=default).parsebytes(
        b'MIME-Version: 1.0\r\nContent-Type: ' + content_type.encode() + b'\r\n\r\n' + body)
    parts = {}
    for part in message.iter_parts():
        name = part.get_param('name', header='content-disposition')
        parts.setdefault(name, []).append(part)
    if not parts.get('project'):
        raise ValueError('Choose a .zip project or .tex file.')
    project = parts['project'][0]
    filename = Path(project.get_filename() or '').name
    if not filename.lower().endswith(('.zip', '.tex')):
        raise ValueError('Choose a .zip project or .tex file.')
    files = {}
    if filename.lower().endswith('.zip'):
        with zipfile.ZipFile(io.BytesIO(project.get_payload(decode=True))) as archive:
            infos = [info for info in archive.infolist() if not info.is_dir()]
            if len(infos) > MAX_FILES or sum(info.file_size for info in infos) > MAX_UNPACKED:
                raise ValueError('The ZIP project is too large to convert.')
            for info in infos:
                name = info.filename.replace('\\', '/')
                path = PurePosixPath(name)
                if '..' in path.parts or path.is_absolute() or path.suffix.lower() not in PROJECT_EXTENSIONS:
                    continue
                files[str(path)] = archive.read(info)
    else:
        files[filename] = project.get_payload(decode=True)
        for field in ('bibliography', 'figures'):
            for part in parts.get(field, []):
                name = Path(part.get_filename() or '').name
                if name and Path(name).suffix.lower() in PROJECT_EXTENSIONS:
                    files[name] = part.get_payload(decode=True)
    if not any(path.lower().endswith('.tex') for path in files):
        raise ValueError('No .tex manuscript was found in the project.')
    selected = parts.get('main', [None])[0]
    main = selected.get_content().strip() if selected else ''
    if main and main not in files:
        raise ValueError('The selected main .tex file is missing from the project.')
    return files, main


def project_summary(files):
    tex = sorted(path for path in files if path.lower().endswith('.tex'))
    bib = sorted(path for path in files if path.lower().endswith('.bib'))
    images = sorted(path for path in files if PurePosixPath(path).suffix.lower() in IMAGE_EXTENSIONS)
    suggested = next((path for path in tex if PurePosixPath(path).name.lower() == 'main.tex'), tex[0])
    return {'tex_files': tex, 'bibliographies': bib, 'figures': images,
            'suggested_main': suggested, 'total_files': len(files)}


class Handler(BaseHTTPRequestHandler):
    def json_response(self, status, data):
        payload = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if urlparse(self.path).path not in ('/', '/index.html'):
            self.send_error(404)
            return
        data = (ROOT / 'index.html').read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        route = urlparse(self.path).path
        if route not in ('/inspect', '/convert'):
            self.send_error(404)
            return
        size = int(self.headers.get('Content-Length', '0'))
        if size <= 0 or size > MAX_UPLOAD:
            self.json_response(413, {'error': 'Upload must be below 40 MB.'})
            return
        try:
            files, main = read_upload(self.headers, self.rfile.read(size))
            if route == '/inspect':
                self.json_response(200, project_summary(files))
                return
            document, report = convert(files, main)
            out = io.BytesIO()
            document.save(out)
            payload = out.getvalue()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', 'attachment; filename="converted-manuscript.docx"')
            self.send_header('X-Conversion-Report', json.dumps(report, separators=(',', ':')))
            self.send_header('Content-Length', str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
        except (ValueError, zipfile.BadZipFile, KeyError) as exc:
            self.json_response(400, {'error': str(exc)})
        except Exception as exc:
            self.json_response(500, {'error': f'Conversion failed: {exc}'})


if __name__ == '__main__':
    print('LaTeX to Word: http://localhost:8765')
    ThreadingHTTPServer(('127.0.0.1', 8765), Handler).serve_forever()
