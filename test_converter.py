import io
import base64
import unittest
import zipfile
import fitz

from converter import convert


class ConverterTests(unittest.TestCase):
    def test_nested_author_and_affiliations_are_formatted(self):
        source = r'''\title{Sample Manuscript}
\author{Ahmad Nawaz Zaheer, \textbf{Muhammad Farhan}, and Awais Ahmad
\thanks{Ahmad Nawaz Zaheer is with Boya International College (e-mail: alpha@example.com).}
\thanks{Muhammad Farhan is with COMSATS University Islamabad.}}
\markboth{Journal Name}{Ahmad \MakeLowercase{et al.}: Sample Manuscript}
\begin{document}\maketitle
Body text.
\end{document}'''
        document, _ = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(document.paragraphs[0].text, 'Sample Manuscript')
        self.assertEqual(document.paragraphs[1].text,
                         'Ahmad Nawaz Zaheer, Muhammad Farhan, and Awais Ahmad')
        self.assertTrue(any(run.bold for run in document.paragraphs[1].runs))
        self.assertIn('Boya International College', document.paragraphs[2].text)
        self.assertIn('COMSATS University Islamabad', document.paragraphs[3].text)
        for raw in (r'\author', r'\thanks', r'\markboth', r'\MakeLowercase', r'\maketitle'):
            self.assertNotIn(raw, xml)

    def test_ieee_author_blocks_render_names_and_affiliations(self):
        source = r'''\author{\IEEEauthorblockN{First Author}\IEEEauthorblockA{First University}
\and \IEEEauthorblockN{Second Author}\IEEEauthorblockA{Second University}}
\begin{document}\maketitle\end{document}'''
        document, _ = convert({'main.tex': source.encode()})
        self.assertEqual(document.paragraphs[0].text, 'First Author, Second Author')
        self.assertEqual(document.paragraphs[1].text, 'First University')
        self.assertEqual(document.paragraphs[2].text, 'Second University')

    def test_math_formatting_and_inline_bibliography(self):
        source = r"""\begin{document}
\section{Introduction}
This is \textbf{important}: $x^2+\frac{1}{2}$ \cite{smith}.
\begin{equation}\sqrt{x}=\alpha\end{equation}
\begin{itemize}\item First\item Second\end{itemize}
\begin{thebibliography}{9}
\bibitem{smith} Smith, J. 2024. A Study.
\end{thebibliography}
\end{document}"""
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['equations'], 2)
        self.assertEqual(report['bibitems'], 1)
        self.assertEqual(report['citations'], 1)
        self.assertIn('<m:f>', xml)
        self.assertIn('<m:rad>', xml)
        self.assertIn('<m:sSup>', xml)
        self.assertIn('ADDIN EN.CITE', xml)
        self.assertIn('ADDIN EN.REFLIST', xml)
        self.assertIn('<w:b', xml)
        self.assertIn('ListBullet', xml)

    def test_equation_labels_become_numbered_bookmarks_and_references(self):
        source = r"""\begin{document}
As shown in \eqref{eq:energy}, and equation \ref{eq:second}.
\begin{equation}E=mc^2\label{eq:energy}\end{equation}
\begin{align}a=b\label{eq:second}\\ c=d\nonumber\end{align}
\end{document}"""
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['equations'], 3)
        self.assertEqual(report['unresolved_references'], [])
        self.assertIn('Equation_1', xml)
        self.assertIn('Equation_2', xml)
        self.assertIn('REF Equation_1', xml)
        self.assertIn('REF Equation_2', xml)
        self.assertNotIn('eq:energy', xml)
        self.assertNotIn('eq:second', xml)

    def test_figure_table_cross_references_and_embedded_image(self):
        png = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL/nwAAAABJRU5ErkJggg==')
        source = r"""\begin{document}
See Figure \ref{fig:one} and \autoref{tab:one}.
\begin{figure}\includegraphics{plot.png}\caption{A plot}\label{fig:one}\end{figure}
\begin{table}\caption{Values}\label{tab:one}\begin{tabular}{cc}A & B \\ 1 & 2\end{tabular}\end{table}
\end{document}"""
        document, report = convert({'main.tex': source.encode(), 'plot.png': png})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
            media = [name for name in archive.namelist() if name.startswith('word/media/')]
        self.assertEqual(report['figures'], 1)
        self.assertEqual(report['tables'], 1)
        self.assertEqual(report['unresolved_references'], [])
        self.assertEqual(len(media), 1)
        self.assertIn('SEQ Figure', xml)
        self.assertIn('SEQ Table', xml)
        self.assertIn('REF Figure_1', xml)
        self.assertIn('REF Table_1', xml)

    def test_math_scripts_attach_to_the_symbol(self):
        document, _ = convert({'main.tex': rb'\begin{document}$x^2+\hat{y}+\binom{n}{k}$\end{document}'})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertIn('<m:sSup>', xml)
        self.assertIn('<m:acc>', xml)
        self.assertIn('<m:f>', xml)

    def test_matrix_delimiters_indexed_root_and_combined_scripts(self):
        source = r'\begin{document}$x_i^2+\sqrt[3]{x}+\left(\frac{a}{b}\right)+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\sum_{i=1}^{n}x_i$\end{document}'
        document, _ = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertIn('<m:sSubSup>', xml)
        self.assertIn('<m:rad>', xml)
        self.assertIn('<m:deg>', xml)
        self.assertIn('<m:d>', xml)
        self.assertIn('<m:m>', xml)
        self.assertIn('<m:nary>', xml)
        self.assertEqual(xml.count('<m:mr>'), 2)

    def test_dollar_math_and_table_caption_order(self):
        source = r'''\begin{document}Cost is \$5 and $x^2$ is math.
$$\frac{a}{b}$$
\begin{table}\begin{tabular}{cc}A & B\\1 & 2\end{tabular}\caption{Values}\end{table}
\end{document}'''
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['equations'], 2)
        self.assertIn('$5', document.paragraphs[0].text)
        self.assertEqual(xml.count('<m:oMath>'), 2)
        self.assertLess(xml.index('SEQ Table'), xml.index('<w:tbl>'))

    def test_png_jpeg_and_pdf_figures_are_embedded(self):
        pixmap = fitz.Pixmap(fitz.csRGB, (0, 0, 12, 12), 0)
        pdf = fitz.open()
        pdf.new_page(width=72, height=72)
        source = r'''\begin{document}
\begin{figure}\includegraphics{one.png}\caption{PNG}\end{figure}
\begin{figure}\includegraphics{two.jpg}\caption{JPEG}\end{figure}
\begin{figure}\includegraphics{three.pdf}\caption{PDF}\end{figure}
\end{document}'''
        files = {'main.tex': source.encode(), 'one.png': pixmap.tobytes('png'),
                 'two.jpg': pixmap.tobytes('jpg'), 'three.pdf': pdf.tobytes()}
        document, report = convert(files)
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            media = [name for name in archive.namelist() if name.startswith('word/media/')]
        self.assertEqual(report['figures'], 3)
        self.assertEqual(report['missing_figures'], [])
        self.assertEqual(len(media), 3)

    def test_table_cell_math_and_algorithm_steps(self):
        source = r'''\begin{document}
See \autoref{alg:solve}.
\begin{table}\caption{Math cells}\begin{tabular}{cc}$x^2$ & $\frac{1}{2}$ \\ $\alpha$ & value\end{tabular}\end{table}
\begin{algorithm}\caption{Solve}\label{alg:solve}
\begin{algorithmic}[1]
\Require $x>0$
\State $y\gets x^2$
\If{$y>1$}
\Return $y$
\EndIf
\end{algorithmic}\end{algorithm}
\end{document}'''
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['tables'], 1)
        self.assertEqual(report['algorithms'], 1)
        self.assertGreaterEqual(report['equations'], 5)
        self.assertIn('<m:sSup>', xml)
        self.assertIn('<m:f>', xml)
        self.assertIn('SEQ Algorithm', xml)
        self.assertIn('REF Algorithm_1', xml)
        self.assertIn('<w:b', xml)
        self.assertEqual(report['unresolved_references'], [])

    def test_manuscript_commands_and_layout_fragments_do_not_leak(self):
        source = r'''\title{Secure and Scalable Intrusion Detection}
%Authors' placeholder
\maketitle
\begin{document}
\begin{IEEEkeywords}\end{IEEEkeywords}
\section{Introduction}\label{sec:introduction}
p{0.35\textwidth}\
p{0.115\textwidth} p{0.35\textwidth}}
\begin{equation}\begin{aligned}a &= b \\ c &= d\end{aligned}\end{equation}
\begin{table}\caption{Values}\begin{tabularx}{\textwidth}{p{0.35\textwidth}p{0.35\textwidth}}$x^2$ & $\mathbb{R}$\\A & B\end{tabularx}\end{table}
\end{document}'''
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['tables'], 1)
        self.assertIn('<m:eqArr>', xml)
        self.assertIn('double-struck', xml)
        for raw in ('\\title', '\\maketitle', 'IEEEkeywords', '\\label', 'textwidth', 'p{0.35', '\\begin{aligned}', '\\end{aligned}'):
            self.assertNotIn(raw, xml)

    def test_graphicspath_and_standalone_image_are_inserted(self):
        pixmap = fitz.Pixmap(fitz.csRGB, (0, 0, 12, 12), 0)
        source = r'''\graphicspath{{figures/}}
\begin{document}
\includegraphics[width=.5\textwidth]{overview}
\begin{figure}\includegraphics{detail.jpg}\caption{Detail}\end{figure}
\end{document}'''
        files = {'paper/main.tex': source.encode(),
                 'paper/figures/overview.png': pixmap.tobytes('png'),
                 'paper/figures/detail.jpg': pixmap.tobytes('jpg')}
        document, report = convert(files)
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            media = [name for name in archive.namelist() if name.startswith('word/media/')]
        self.assertEqual(report['figures'], 2)
        self.assertEqual(report['missing_figures'], [])
        self.assertEqual(len(media), 2)

    def test_nested_bold_shortstack_and_grouped_math(self):
        source = r'''\begin{document}
This is \textbf{bold with \textit{italic inside}}.
\begin{table}\begin{tabular}{cc}
\shortstack[l]{\textbf{First}\\Second line} & $\mathbf{x}+{a+b}^2$ \\
Value & $\boldsymbol{\alpha}$
\end{tabular}\end{table}
\end{document}'''
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO()
        document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['tables'], 1)
        self.assertEqual(len(document.tables[0].rows), 2)
        self.assertEqual(document.tables[0].cell(0, 0).text, 'First\nSecond line')
        self.assertIn('<w:br', xml)
        self.assertIn('<m:box>', xml)
        self.assertIn('<m:sSup>', xml)
        self.assertIn('m:val="b"', xml)
        self.assertTrue(any(run.bold for run in document.paragraphs[0].runs))
        self.assertTrue(any(run.bold and run.italic for run in document.paragraphs[0].runs))

    def test_paper_math_spacing_delimiters_and_large_operator(self):
        source = r'''\begin{document}
\begin{equation}
U_o=\bigoplus_{j=1}^{G_o}\left(\beta_j+\sum_{k=1}^{G_d}\delta_{jk}F_k\right),\qquad
\left\|\Delta\theta_n\right\|_2+\widetilde{E}_n+\mathbf{1}\!\left(x\right)\wedge\top.
\end{equation}\end{document}'''
        document, _ = convert({'main.tex': source.encode()})
        output = io.BytesIO(); document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertIn('<m:nary>', xml)
        self.assertIn('⨁', xml)
        self.assertIn('∧', xml)
        self.assertIn('⊤', xml)
        self.assertIn('m:val="‖"', xml)
        self.assertIn('<m:acc>', xml)
        for raw in ('\\bigoplus', '\\qquad', '\\!', '\\wedge', '\\top', '\\widetilde'):
            self.assertNotIn(raw, xml)

    def test_cases_nested_inside_aligned_equation_preserves_label(self):
        source = r'''\begin{document}See \eqref{eq:policy}.
\begin{equation}
\begin{aligned}
P_n &= \lambda_v(1-v_n), \\
\mathrm{Slash}_n &= \begin{cases}
\mathrm{RevokeAccess}(n), & P_n\geq\tau, \\
\mathrm{NoAction}, & P_n<\tau.
\end{cases}
\end{aligned}
\label{eq:policy}
\end{equation}\end{document}'''
        document, report = convert({'main.tex': source.encode()})
        output = io.BytesIO(); document.save(output)
        with zipfile.ZipFile(output) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertEqual(report['unresolved_references'], [])
        self.assertEqual(xml.count('SEQ Equation'), 1)
        self.assertIn('<m:eqArr>', xml)
        self.assertIn('<m:m>', xml)
        self.assertNotIn('\\end{cases}', xml)


if __name__ == '__main__':
    unittest.main()
