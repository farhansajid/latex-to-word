import io
import unittest
import zipfile
from unittest.mock import patch
from http.server import ThreadingHTTPServer
from threading import Thread
from urllib.request import Request, urlopen
import json

from app import Handler, project_summary, read_upload


def multipart(fields):
    boundary = 'manuscript-test-boundary'
    body = bytearray()
    for name, filename, value in fields:
        body.extend(f'--{boundary}\r\n'.encode())
        disposition = f'Content-Disposition: form-data; name="{name}"'
        if filename: disposition += f'; filename="{filename}"'
        body.extend((disposition + '\r\n\r\n').encode())
        body.extend(value)
        body.extend(b'\r\n')
    body.extend(f'--{boundary}--\r\n'.encode())
    return {'Content-Type': f'multipart/form-data; boundary={boundary}'}, bytes(body)


class AppTests(unittest.TestCase):
    def test_zip_inventory_and_main_selection(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w') as archive:
            archive.writestr('paper/main.tex', r'\begin{document}Main\end{document}')
            archive.writestr('paper/appendix.tex', 'Appendix')
            archive.writestr('paper/refs.bib', '')
            archive.writestr('paper/figures/plot.png', b'PNG')
        headers, body = multipart([('project', 'paper.zip', output.getvalue()),
                                   ('main', None, b'paper/appendix.tex')])
        files, main = read_upload(headers, body)
        summary = project_summary(files)
        self.assertEqual(main, 'paper/appendix.tex')
        self.assertEqual(summary['suggested_main'], 'paper/main.tex')
        self.assertEqual(len(summary['tex_files']), 2)
        self.assertEqual(summary['figures'], ['paper/figures/plot.png'])

    def test_zip_size_limit(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('main.tex', b'x' * 1024)
        headers, body = multipart([('project', 'large.zip', output.getvalue())])
        with patch('app.MAX_UNPACKED', 100):
            with self.assertRaisesRegex(ValueError, 'too large'):
                read_upload(headers, body)

    def test_http_inspect_then_convert(self):
        archive_bytes = io.BytesIO()
        with zipfile.ZipFile(archive_bytes, 'w') as archive:
            archive.writestr('main.tex', r'\begin{document}Hello $x^2$.\end{document}')
        headers, body = multipart([('project', 'project.zip', archive_bytes.getvalue())])
        server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        thread = Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            base = f'http://127.0.0.1:{server.server_port}'
            with urlopen(Request(base + '/inspect', body, headers)) as response:
                inventory = json.load(response)
            self.assertEqual(inventory['tex_files'], ['main.tex'])
            with urlopen(Request(base + '/convert', body, headers)) as response:
                document = response.read()
                report = json.loads(response.headers['X-Conversion-Report'])
            self.assertEqual(report['equations'], 1)
            with zipfile.ZipFile(io.BytesIO(document)) as archive:
                self.assertIn('word/document.xml', archive.namelist())
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=2)


if __name__ == '__main__':
    unittest.main()
