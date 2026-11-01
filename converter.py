from __future__ import annotations

import io
import re
from pathlib import PurePosixPath
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH

GREEK = {'alpha':'α','beta':'β','gamma':'γ','delta':'δ','epsilon':'ε','varepsilon':'ε','zeta':'ζ','eta':'η','theta':'θ','vartheta':'ϑ','iota':'ι','kappa':'κ','lambda':'λ','mu':'μ','nu':'ν','xi':'ξ','pi':'π','rho':'ρ','varrho':'ϱ','sigma':'σ','tau':'τ','upsilon':'υ','phi':'φ','varphi':'ϕ','chi':'χ','psi':'ψ','omega':'ω','Gamma':'Γ','Delta':'Δ','Theta':'Θ','Lambda':'Λ','Xi':'Ξ','Pi':'Π','Sigma':'Σ','Phi':'Φ','Psi':'Ψ','Omega':'Ω'}
SYMBOLS = {'times':'×','cdot':'·','pm':'±','mp':'∓','leq':'≤','geq':'≥','neq':'≠','equiv':'≡','infty':'∞','partial':'∂','nabla':'∇','sum':'∑','prod':'∏','int':'∫','iint':'∬','bigoplus':'⨁','bigcup':'⋃','oplus':'⊕','odot':'⊙','wedge':'∧','vee':'∨','top':'⊤','ell':'ℓ','rightarrow':'→','leftarrow':'←','leftrightarrow':'↔','gets':'←','in':'∈','notin':'∉','subset':'⊂','subseteq':'⊆','cup':'∪','cap':'∩','approx':'≈','sim':'∼','propto':'∝','forall':'∀','exists':'∃','ldots':'…','cdots':'⋯','langle':'⟨','rangle':'⟩','Vert':'‖','xi':'ξ','Xi':'Ξ'}


def formatted_run(paragraph, text, styles=()):
    run = paragraph.add_run(text)
    if 'textbf' in styles: run.bold = True
    if 'textit' in styles or 'emph' in styles: run.italic = True
    if 'underline' in styles: run.underline = True
    if 'texttt' in styles: run.font.name = 'Consolas'


def latex_text(paragraph, text, styles=()):
    """Render nested LaTeX text styles as actual Word run properties."""
    command = re.compile(r'\\([A-Za-z]+)(?:\[[^]]*\])?')
    pos, buffer = 0, []
    def flush():
        if buffer:
            formatted_run(paragraph, ''.join(buffer).replace('~', ' '), styles)
            buffer.clear()
    while pos < len(text):
        if text[pos] != '\\':
            buffer.append(text[pos]); pos += 1; continue
        if pos + 1 < len(text) and text[pos + 1] in '&%_$#{}':
            buffer.append(text[pos + 1]); pos += 2; continue
        match = command.match(text, pos)
        if not match:
            buffer.append(text[pos]); pos += 1; continue
        name = match.group(1)
        if name in ('textbf', 'textit', 'emph', 'underline', 'texttt') and match.end() < len(text) and text[match.end()] == '{':
            flush()
            inner, pos = brace_group(text, match.end())
            latex_text(paragraph, inner, styles + (name,))
            continue
        if name in GREEK or name in SYMBOLS:
            flush()
            formatted_run(paragraph, GREEK.get(name, SYMBOLS.get(name)), styles)
            pos = match.end()
            continue
        buffer.append(text[pos:match.end()]); pos = match.end()
    flush()


def math_node(parent, name):
    element = OxmlElement('m:' + name)
    parent.append(element)
    return element


def math_run(parent, value, script=None, bold=False, plain=False):
    run = math_node(parent, 'r')
    if script or bold or plain:
        props = math_node(run, 'rPr')
        if script:
            style = math_node(props, 'scr')
            style.set(qn('m:val'), script)
        if bold:
            style = math_node(props, 'sty')
            style.set(qn('m:val'), 'b')
        elif plain:
            style = math_node(props, 'sty')
            style.set(qn('m:val'), 'p')
    text = math_node(run, 't')
    text.text = value


def math_tokens(source):
    token_re = re.compile(r'\\[A-Za-z]+|\\.|[{}_^]|[A-Za-z]+|\d+(?:\.\d+)?|\s+|.')
    return [m.group() for m in token_re.finditer(source)]


def strip_comments(source):
    return re.sub(r'(?<!\\)%[^\n]*', '', source)


def brace_group(source, start):
    if start >= len(source) or source[start] != '{': return '', start
    depth = 0
    for pos in range(start, len(source)):
        if source[pos] == '{': depth += 1
        elif source[pos] == '}':
            depth -= 1
            if depth == 0: return source[start + 1:pos], pos + 1
    return source[start + 1:], len(source)


def take_command_groups(source, name, count=1):
    """Remove a command and its balanced arguments, returning their contents."""
    pattern = re.compile(r'\\' + re.escape(name) + r'(?![A-Za-z])\s*\{')
    groups, result, pos = [], [], 0
    for match in pattern.finditer(source):
        if match.start() < pos: continue
        values, end = [], match.end() - 1
        for _ in range(count):
            while end < len(source) and source[end].isspace(): end += 1
            if end >= len(source) or source[end] != '{': break
            value, end = brace_group(source, end)
            values.append(value)
        if len(values) != count: continue
        result.append(source[pos:match.start()])
        groups.append(values)
        pos = end
    result.append(source[pos:])
    return ''.join(result), groups


def tabular_body(source):
    match = re.search(r'\\begin\{(tabular|tabularx|longtable)\}', source)
    if not match: return None
    pos = match.end()
    needed = 2 if match.group(1) == 'tabularx' else 1
    for _ in range(needed):
        while pos < len(source) and source[pos].isspace(): pos += 1
        _, pos = brace_group(source, pos)
    end = re.search(r'\\end\{' + match.group(1) + r'\}', source[pos:])
    return source[pos:pos + end.start()] if end else source[pos:]


def split_math_rows(source):
    """Split LaTeX alignment rows without breaking nested cases or matrices."""
    parts = re.findall(r'\\begin\{[^{}]+\}|\\end\{[^{}]+\}|\\\\|[\s\S]', source)
    rows, current, depth = [], [], 0
    for part in parts:
        if part.startswith(r'\begin{'): depth += 1
        elif part.startswith(r'\end{'): depth = max(0, depth - 1)
        if part == r'\\' and depth == 0:
            rows.append(''.join(current))
            current = []
        else:
            current.append(part)
    rows.append(''.join(current))
    return [row.strip() for row in rows if row.strip()]


SHORTSTACK_BREAK = '\ue000'


def protect_shortstacks(source):
    pattern = re.compile(r'\\shortstack(?:\[[lcr]\])?\{')
    result, pos = [], 0
    for match in pattern.finditer(source):
        if match.start() < pos: continue
        inner, end = brace_group(source, match.end() - 1)
        result.append(source[pos:match.end()])
        result.append(inner.replace('\\\\', SHORTSTACK_BREAK))
        result.append('}')
        pos = end
    result.append(source[pos:])
    return ''.join(result)


def table_cell_content(paragraph, value, entries, numbering, report, labels):
    match = re.match(r'\\shortstack(?:\[([lcr])\])?\{', value)
    if match:
        inner, end = brace_group(value, match.end() - 1)
        if not value[end:].strip():
            alignment = {'l': WD_ALIGN_PARAGRAPH.LEFT, 'c': WD_ALIGN_PARAGRAPH.CENTER, 'r': WD_ALIGN_PARAGRAPH.RIGHT}
            paragraph.alignment = alignment.get(match.group(1) or 'c')
            for index, line in enumerate(inner.split(SHORTSTACK_BREAK)):
                if index: paragraph.add_run().add_break()
                add_text(paragraph, line.strip(), entries, numbering, report, labels)
            return
    add_text(paragraph, value.replace(SHORTSTACK_BREAK, '\n'), entries, numbering, report, labels)


def build_math(parent, source):
    """Build editable OMML for common LaTeX math constructs."""
    tokens = math_tokens(source)
    pos = 0
    def group():
        nonlocal pos
        if pos >= len(tokens): return ''
        if tokens[pos] != '{':
            value = tokens[pos]; pos += 1; return value
        pos += 1; depth = 1; result = []
        while pos < len(tokens) and depth:
            token = tokens[pos]; pos += 1
            if token == '{': depth += 1
            elif token == '}': depth -= 1
            if depth: result.append(token)
        return ''.join(result)
    def normalize_delimiter(value):
        return {'\\{': '{', '\\}': '}', '\\|': '‖', '\\Vert': '‖',
                '\\langle': '⟨', '\\rangle': '⟩', '\\[': '[', '\\]': ']'}.get(value, value)
    def nary_operand():
        nonlocal pos
        while pos < len(tokens) and tokens[pos].isspace(): pos += 1
        result, depth, delimiter_depth = [], 0, 0
        while pos < len(tokens):
            current = tokens[pos]
            if depth == 0 and delimiter_depth == 0 and current in ('+', '-', '=', ',', ';', r'\right', r'\\'):
                break
            if current == '{': depth += 1
            elif current == '}': depth = max(0, depth - 1)
            elif current == r'\left': delimiter_depth += 1
            elif current == r'\right': delimiter_depth = max(0, delimiter_depth - 1)
            result.append(current); pos += 1
        return ''.join(result).strip()
    while pos < len(tokens):
        token = tokens[pos]; pos += 1
        if token.isspace(): continue
        if token == r'\left':
            opening = normalize_delimiter(tokens[pos]) if pos < len(tokens) else '('; pos += 1
            depth, contents, closing = 1, [], ')'
            while pos < len(tokens):
                current = tokens[pos]; pos += 1
                if current == r'\left': depth += 1
                elif current == r'\right':
                    depth -= 1
                    if depth == 0:
                        closing = normalize_delimiter(tokens[pos]) if pos < len(tokens) else ')'; pos += 1
                        break
                contents.append(current)
            delimiter = math_node(parent, 'd')
            props = math_node(delimiter, 'dPr')
            for name, value in [('begChr', opening), ('endChr', closing)]:
                setting = math_node(props, name)
                setting.set(qn('m:val'), '' if value == '.' else value)
            build_math(math_node(delimiter, 'e'), ''.join(contents))
        elif token == r'\begin':
            environment = group()
            if environment in ('matrix', 'pmatrix', 'bmatrix', 'Bmatrix', 'vmatrix', 'Vmatrix', 'cases', 'aligned', 'gathered', 'split'):
                contents = []
                ending = [r'\end', '{', environment, '}']
                while pos < len(tokens):
                    if tokens[pos:pos + 4] == ending:
                        pos += 4
                        break
                    contents.append(tokens[pos]); pos += 1
                outer = parent
                if environment in ('aligned', 'gathered', 'split'):
                    arrangement = math_node(parent, 'eqArr')
                    for row in split_math_rows(''.join(contents)):
                        if row.strip(): build_math(math_node(arrangement, 'e'), row.replace('&', '').strip())
                    continue
                if environment != 'matrix':
                    chars = {'pmatrix':('(',')'),'bmatrix':('[',']'),'Bmatrix':('{','}'),'vmatrix':('|','|'),'Vmatrix':('‖','‖'),'cases':('{','')}
                    outer = math_node(parent, 'd')
                    props = math_node(outer, 'dPr')
                    for name, value in zip(('begChr', 'endChr'), chars[environment]):
                        setting = math_node(props, name); setting.set(qn('m:val'), value)
                    outer = math_node(outer, 'e')
                matrix = math_node(outer, 'm')
                for row in split_math_rows(''.join(contents)):
                    if not row.strip(): continue
                    row_node = math_node(matrix, 'mr')
                    for cell in row.split('&'):
                        build_math(math_node(row_node, 'e'), cell.strip())
            else:
                math_run(parent, token + '{' + environment + '}')
        elif token in (r'\frac', r'\dfrac', r'\tfrac', r'\binom'):
            numerator, denominator = group(), group()
            frac = math_node(parent, 'f')
            if token == r'\binom':
                props = math_node(frac, 'fPr')
                ftype = math_node(props, 'type')
                ftype.set(qn('m:val'), 'noBar')
            build_math(math_node(frac, 'num'), numerator)
            build_math(math_node(frac, 'den'), denominator)
        elif token == r'\sqrt':
            rad = math_node(parent, 'rad')
            degree = ''
            if pos < len(tokens) and tokens[pos] == '[':
                pos += 1
                while pos < len(tokens) and tokens[pos] != ']':
                    degree += tokens[pos]; pos += 1
                if pos < len(tokens): pos += 1
            if degree: build_math(math_node(rad, 'deg'), degree)
            else:
                props = math_node(rad, 'radPr')
                hidden = math_node(props, 'degHide')
                hidden.set(qn('m:val'), '1')
                math_node(rad, 'deg')
            build_math(math_node(rad, 'e'), group())
        elif token in (r'\overline', r'\bar', r'\hat', r'\vec', r'\tilde', r'\widetilde'):
            accent = math_node(parent, 'acc')
            props = math_node(accent, 'accPr')
            mark = math_node(props, 'chr')
            mark.set(qn('m:val'), {'\\overline':'¯','\\bar':'¯','\\hat':'̂','\\vec':'⃗','\\tilde':'̃','\\widetilde':'̃'}[token])
            build_math(math_node(accent, 'e'), group())
        elif token in (r'\overbrace', r'\underbrace'):
            brace = math_node(parent, 'groupChr')
            props = math_node(brace, 'groupChrPr')
            mark = math_node(props, 'chr')
            mark.set(qn('m:val'), '⏞' if token == r'\overbrace' else '⏟')
            position = math_node(props, 'pos')
            position.set(qn('m:val'), 'top' if token == r'\overbrace' else 'bot')
            build_math(math_node(brace, 'e'), group())
        elif token in (r'\mathbb', r'\mathcal', r'\mathfrak'):
            style = {r'\mathbb':'double-struck',r'\mathcal':'script',r'\mathfrak':'fraktur'}[token]
            math_run(parent, group(), style)
        elif token in (r'\mathrm', r'\mathbf', r'\mathit', r'\text', r'\operatorname', r'\boldsymbol', r'\textbf'):
            value = group()
            rendered = re.sub(r'\\([A-Za-z]+)', lambda m: GREEK.get(m.group(1), SYMBOLS.get(m.group(1), m.group(0))), value)
            math_run(parent, rendered, bold=token in (r'\mathbf', r'\boldsymbol', r'\textbf'),
                     plain=token in (r'\mathrm', r'\text', r'\operatorname'))
        elif token in (r'\sum', r'\prod', r'\int', r'\iint', r'\bigoplus', r'\bigcup'):
            command = token[1:]
            lower, upper = '', ''
            while pos < len(tokens) and tokens[pos] in ('_', '^'):
                marker = tokens[pos]; pos += 1
                if marker == '_': lower = group()
                else: upper = group()
            nary = math_node(parent, 'nary')
            props = math_node(nary, 'naryPr')
            char = math_node(props, 'chr'); char.set(qn('m:val'), SYMBOLS[command])
            build_math(math_node(nary, 'sub'), lower)
            build_math(math_node(nary, 'sup'), upper)
            operand = nary_operand()
            if operand: build_math(math_node(nary, 'e'), operand)
            else: math_node(nary, 'e')
        elif token in ('^', '_'):
            value = group()
            if len(parent) and parent[-1].tag in (qn('m:sSup'), qn('m:sSub')):
                previous = parent[-1]
                opposite = (token == '^' and previous.tag == qn('m:sSub')) or (token == '_' and previous.tag == qn('m:sSup'))
                if opposite:
                    parent.remove(previous)
                    combined = math_node(parent, 'sSubSup')
                    combined.append(previous.find(qn('m:e')))
                    if token == '^':
                        combined.append(previous.find(qn('m:sub')))
                        build_math(math_node(combined, 'sup'), value)
                    else:
                        build_math(math_node(combined, 'sub'), value)
                        combined.append(previous.find(qn('m:sup')))
                    continue
            script = math_node(parent, 'sSup' if token == '^' else 'sSub')
            base = math_node(script, 'e')
            if len(parent) > 1:
                previous = parent[-2]
                parent.remove(previous)
                base.append(previous)
            build_math(math_node(script, 'sup' if token == '^' else 'sub'), value)
        elif token == '{':
            depth, contents = 1, []
            while pos < len(tokens) and depth:
                current = tokens[pos]; pos += 1
                if current == '{': depth += 1
                elif current == '}': depth -= 1
                if depth: contents.append(current)
            box = math_node(parent, 'box')
            build_math(math_node(box, 'e'), ''.join(contents))
        elif token.startswith('\\'):
            command = token[1:]
            if command in ('left', 'right'): continue
            if command in ('big', 'Big', 'bigg', 'Bigg', 'bigl', 'bigr', 'Bigl', 'Bigr', 'biggl', 'biggr', 'Biggl', 'Biggr'):
                continue
            if command in ('quad', 'qquad'):
                math_run(parent, '\u2003' * (2 if command == 'qquad' else 1), plain=True)
                continue
            if command in ('sin', 'cos', 'tan', 'tanh', 'log', 'ln', 'exp', 'lim', 'max', 'min'):
                math_run(parent, command, plain=True)
                continue
            if token in (r'\!', r'\,', r'\:', r'\;'):
                if token != r'\!': math_run(parent, '\u2009', plain=True)
                continue
            if token in (r'\{', r'\}', r'\|'):
                math_run(parent, normalize_delimiter(token), plain=True)
                continue
            math_run(parent, GREEK.get(command, SYMBOLS.get(command, token)))
        elif token not in ('{', '}'):
            math_run(parent, token)


def add_equation(paragraph, source):
    math = OxmlElement('m:oMath')
    build_math(math, source.strip())
    paragraph._p.append(math)


def equation_lines(token):
    body = re.sub(r'^\\begin\{[^}]+\}|\\end\{[^}]+\}$', '', token.strip())
    if body.startswith(r'\['): body = body[2:-2]
    if body.startswith('$$') and body.endswith('$$'): body = body[2:-2]
    return split_math_rows(body)


def equation_numbered(token, line):
    return (token.startswith(r'\begin{equation}') or token.startswith(r'\begin{align}')) and not re.search(r'\\(?:nonumber|notag)\b', line)


def equation_labels(tokens):
    labels, numbers = {}, {}
    count = 0
    for index, token in enumerate(tokens):
        if not (token.startswith(r'\begin{equation') or token.startswith(r'\begin{align') or token.startswith(r'\begin{displaymath') or token.startswith(r'\[')):
            continue
        for line_index, line in enumerate(equation_lines(token)):
            if equation_numbered(token, line):
                count += 1
                numbers[(index, line_index)] = count
                for label in re.findall(r'\\label\{([^}]+)\}', line):
                    labels[label] = count
    return labels, numbers


def bookmark_number(paragraph, number):
    paragraph.add_run('(')
    marker = OxmlElement('w:bookmarkStart')
    marker.set(qn('w:id'), str(number))
    marker.set(qn('w:name'), f'Equation_{number}')
    paragraph._p.append(marker)
    field(paragraph, ' SEQ Equation \\* ARABIC ', str(number))
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(number))
    paragraph._p.append(end)
    paragraph.add_run(')')


def caption_number(paragraph, kind, number):
    paragraph.add_run(kind + ' ')
    marker = OxmlElement('w:bookmarkStart')
    marker.set(qn('w:id'), str({'Figure': 1000, 'Table': 2000, 'Algorithm': 3000}[kind] + number))
    marker.set(qn('w:name'), f'{kind}_{number}')
    paragraph._p.append(marker)
    field(paragraph, f' SEQ {kind} \\* ARABIC ', str(number))
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str({'Figure': 1000, 'Table': 2000, 'Algorithm': 3000}[kind] + number))
    paragraph._p.append(end)
    paragraph.add_run('. ')


def asset_labels(tokens):
    labels, counts = {}, {'Figure': 0, 'Table': 0, 'Algorithm': 0}
    numbers = {}
    for index, token in enumerate(tokens):
        kind = ('Figure' if token.startswith(r'\begin{figure') else
                'Table' if token.startswith(r'\begin{table') else
                'Algorithm' if token.startswith(r'\begin{algorithm') else None)
        if not kind: continue
        counts[kind] += 1
        numbers[index] = (kind, counts[kind])
        for label in re.findall(r'\\label\{([^}]+)\}', token):
            labels[label] = (kind, counts[kind])
    return labels, numbers


def inline_bibitems(source):
    match = re.search(r'\\begin\{thebibliography\}(?:\{[^}]*\})?(.*?)\\end\{thebibliography\}', source, re.S)
    if not match: return source, {}
    entries = {}
    for item in re.finditer(r'\\bibitem(?:\[[^]]*\])?\{([^}]+)\}(.*?)(?=\\bibitem|$)', match.group(1), re.S):
        value = clean(re.sub(r'\s+', ' ', item.group(2)))
        year = re.search(r'\b(?:19|20)\d{2}\b', value)
        entries[item.group(1)] = {'kind':'misc','author':value.split('.')[0], 'title':value,
                                  'year':year.group() if year else '', 'raw':value}
    return source[:match.start()] + source[match.end():], entries


def bib_entries(source):
    entries = {}
    for match in re.finditer(r'@(\w+)\s*\{\s*([^,]+)\s*,', source):
        kind, key = match.group(1), match.group(2).strip()
        pos, depth, quote = match.end(), 1, False
        while pos < len(source) and depth:
            ch = source[pos]
            if ch == '"' and source[pos - 1] != '\\': quote = not quote
            if not quote:
                if ch == '{': depth += 1
                elif ch == '}': depth -= 1
            pos += 1
        chunk = source[match.end():pos - 1]
        fields = {}
        for field in re.finditer(r'(\w+)\s*=\s*(\{(?:[^{}]|\{[^{}]*\})*\}|"[^"]*"|[^,]+)', chunk):
            fields[field.group(1).lower()] = field.group(2).strip().strip('{}" ')
        fields['kind'] = kind.lower()
        entries[key] = fields
    return entries


def clean(text):
    text = re.sub(r'(?<!\\)%[^\n]*', '', text)
    text = re.sub(r'\\(?:emph|textit|textbf|underline)\{([^{}]*)\}', r'\1', text)
    text = re.sub(r'\\(?:label|index)\{[^{}]*\}', '', text)
    text = text.replace('``', '“').replace("''", '”').replace('~', ' ')
    text = re.sub(r'\\([&%_$#{}])', r'\1', text)
    return text.strip()


def add_author_block(doc, source):
    """Render author names and affiliation notes from a LaTeX author argument."""
    source, thanks = take_command_groups(source, 'thanks')
    source, affiliations = take_command_groups(source, 'IEEEauthorblockA')
    source, extra_affiliations = take_command_groups(source, 'affil')
    source, name_blocks = take_command_groups(source, 'IEEEauthorblockN')
    names = ', '.join(group[0] for group in name_blocks) if name_blocks else source
    names = re.sub(r'\\and\b|\\\\', ', ', names)
    names = re.sub(r'\s*,\s*', ', ', names)
    names = re.sub(r'(?:,\s*){2,}', ', ', names)
    names = re.sub(r'\s+', ' ', names).strip(' ,')
    if names:
        paragraph = doc.add_paragraph(style='Subtitle')
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        latex_text(paragraph, names)
    for note in thanks + affiliations + extra_affiliations:
        value = re.sub(r'\\\\', '; ', note[0])
        value = re.sub(r'\s+', ' ', value).strip()
        if value:
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_after = 0
            latex_text(paragraph, value)
            for run in paragraph.runs: run.font.size = Pt(9)


def field(paragraph, instruction, display):
    run = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    code = OxmlElement('w:instrText'); code.set(qn('xml:space'), 'preserve'); code.text = instruction
    separate = OxmlElement('w:fldChar'); separate.set(qn('w:fldCharType'), 'separate')
    result = OxmlElement('w:t'); result.text = display
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    for element in (begin, code, separate, result, end): run._r.append(element)


def endnote_record(key, entry, number):
    authors = [a.strip() for a in re.split(r'\s+and\s+', entry.get('author', '')) if a.strip()]
    author_xml = ''.join(f'<author>{escape(a)}</author>' for a in authors)
    kind = entry.get('kind', 'article')
    type_name, type_id = ('Book', 6) if kind == 'book' else ('Conference Paper', 10) if kind in ('inproceedings', 'conference') else ('Journal Article', 17)
    tags = [f'<rec-number>{number}</rec-number>', f'<ref-type name="{type_name}">{type_id}</ref-type>',
            f'<contributors><authors>{author_xml}</authors></contributors>',
            f'<titles><title>{escape(entry.get("title", key))}</title><secondary-title>{escape(entry.get("journal", entry.get("booktitle", "")))}</secondary-title></titles>',
            f'<dates><year>{escape(entry.get("year", ""))}</year></dates>']
    for source, target in [('volume', 'volume'), ('number', 'number'), ('pages', 'pages'), ('doi', 'electronic-resource-num'), ('publisher', 'publisher')]:
        if entry.get(source): tags.append(f'<{target}>{escape(entry[source])}</{target}>')
    return ''.join(tags)


def add_text(paragraph, text, entries, numbering, report, labels=None):
    labels = labels or {}
    text = re.sub(r'\\(?:label|index)\{[^}]*\}', '', text)
    pattern = r'(\\(?:cite|citep|citet|citeauthor|citeyear)\*?(?:\[[^]]*\])?\{[^}]+\}|\\(?:eqref|ref|autoref|cref|Cref)\{[^}]+\}|(?<!\\)\$(?!\$)[^$]+(?<!\\)\$|\\\([^)]*\\\))'
    for part in re.split(pattern, text):
        reference = re.fullmatch(r'\\(eqref|ref|autoref|cref|Cref)\{([^}]+)\}', part)
        if reference:
            key = reference.group(2)
            if key in labels:
                kind, number = labels[key]
                if reference.group(1) in ('autoref', 'cref', 'Cref'):
                    paragraph.add_run(('Equation' if kind == 'Equation' else kind) + ' ')
                if reference.group(1) == 'eqref': paragraph.add_run('(')
                field(paragraph, f' REF {kind}_{number} \\h ', str(number))
                if reference.group(1) == 'eqref': paragraph.add_run(')')
            else:
                report['unresolved_references'].append(key)
                paragraph.add_run(f'[?{key}]')
            continue
        if (part.startswith('$') and part.endswith('$')) or (part.startswith(r'\(') and part.endswith(r'\)')):
            add_equation(paragraph, part[1:-1] if part.startswith('$') else part[2:-2])
            report['equations'] += 1
            continue
        cite = re.fullmatch(r'\\(cite|citep|citet|citeauthor|citeyear)\*?(?:\[([^]]*)\])?\{([^}]+)\}', part)
        if not cite:
            latex_text(paragraph, part)
            continue
        keys = [k.strip() for k in cite.group(3).split(',')]
        for key in keys:
            if key not in entries:
                report['unresolved_citations'].append(key)
                paragraph.add_run(f'[?{key}]')
                continue
            entry = entries[key]
            number = numbering.setdefault(key, len(numbering) + 1)
            surname = entry.get('author', key).split(' and ')[0].split(',')[0].split()[-1]
            year = entry.get('year', '')
            display = f'[{number}]'
            payload = (f'<EndNote><Cite><Author>{escape(surname)}</Author><Year>{escape(year)}</Year>'
                       f'<RecNum>{number}</RecNum><DisplayText>{display}</DisplayText><record>'
                       f'{endnote_record(key, entry, number)}</record></Cite></EndNote>')
            field(paragraph, ' ADDIN EN.CITE ' + payload, display)
            report['citations'] += 1


def resolve_image(name, files, tex_path, graphic_paths=()):
    name = name.strip().replace('\\', '/')
    root = PurePosixPath(tex_path).parent
    requested = [str(root / name), name]
    requested += [str(root / prefix / name) for prefix in graphic_paths]
    requested += [str(PurePosixPath(prefix) / name) for prefix in graphic_paths]
    candidates = []
    for value in requested:
        value = value.removeprefix('./').lower()
        candidates.append(value)
        if not PurePosixPath(value).suffix:
            candidates.extend(value + ext for ext in ('.png', '.jpg', '.jpeg', '.pdf'))
    normalized = {path.replace('\\', '/').lower().removeprefix('./'): path for path in files}
    for candidate in candidates:
        if candidate in normalized: return normalized[candidate]
    for candidate in candidates:
        matches = [original for path, original in normalized.items() if path.endswith('/' + candidate)]
        if len(matches) == 1: return matches[0]
    basename = PurePosixPath(name).name.lower()
    matches = [original for path, original in normalized.items() if PurePosixPath(path).name == basename or
               (not PurePosixPath(basename).suffix and PurePosixPath(path).stem == basename and PurePosixPath(path).suffix in ('.png', '.jpg', '.jpeg', '.pdf'))]
    return matches[0] if len(matches) == 1 else None


def insert_figure(doc, name, files, main, graphic_paths, report):
    path = resolve_image(name, files, main, graphic_paths)
    if not path:
        report['missing_figures'].append(name)
        return
    try:
        image = files[path]
        if path.lower().endswith('.pdf'):
            import fitz
            pdf = fitz.open(stream=image, filetype='pdf')
            image = pdf[0].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).tobytes('png')
        doc.add_picture(io.BytesIO(image), width=Inches(5.5))
        report['figures'] += 1
    except Exception as exc:
        report['missing_figures'].append(f'{name} ({type(exc).__name__})')


def convert(files, main=''):
    tex_files = [p for p in files if p.lower().endswith('.tex')]
    if not tex_files: raise ValueError('No .tex file found.')
    main = main if main in files else next((p for p in tex_files if PurePosixPath(p).name.lower() == 'main.tex'), tex_files[0])
    source = files[main].decode('utf-8-sig', errors='replace')
    for path in tex_files:
        if path == main: continue
        source = re.sub(r'\\(?:input|include)\{([^}]+)\}', lambda m: files[path].decode('utf-8-sig', errors='replace') if PurePosixPath(path).stem == PurePosixPath(m.group(1)).stem else m.group(0), source)
    source = strip_comments(source)
    graphic_paths = []
    for specification in re.findall(r'\\graphicspath\{((?:\{[^{}]*\})+)\}', source):
        graphic_paths.extend(re.findall(r'\{([^{}]*)\}', specification))
    source = re.sub(r'\\graphicspath\{(?:\{[^{}]*\})+\}', '', source)
    source, inline_entries = inline_bibitems(source)
    entries = dict(inline_entries)
    for path, data in files.items():
        if path.lower().endswith('.bib'): entries.update(bib_entries(data.decode('utf-8-sig', errors='replace')))
    report = {'main': main, 'citations': 0, 'figures': 0, 'tables': 0, 'algorithms': 0, 'equations': 0, 'bibitems': len(inline_entries), 'unresolved_citations': [], 'unresolved_references': [], 'missing_figures': [], 'warnings': []}
    doc = Document(); numbering = {}
    source, titles = take_command_groups(source, 'title')
    if titles:
        heading = doc.add_heading(level=0)
        latex_text(heading, titles[0][0])
    source, authors = take_command_groups(source, 'author')
    if authors: add_author_block(doc, authors[0][0])
    source, _ = take_command_groups(source, 'date')
    source, _ = take_command_groups(source, 'markboth', 2)
    source = re.sub(r'\\maketitle\b', '', source)
    source = re.sub(r'(?m)^[ \t{}\\|]*(?:[pmb]\{[^{}\n]*\}[ \t{}\\|]*)+$', '', source)
    source = re.sub(r'^.*?\\begin\{document\}', '', source, count=1, flags=re.S) if '\\begin{document}' in source else source
    source = source.split('\\end{document}')[0]
    environments = ('figure', 'figure*', 'table', 'table*', 'algorithm', 'algorithm*', 'algorithmic',
                    'tabular', 'tabularx', 'longtable', 'equation', 'equation*', 'align', 'align*',
                    'displaymath', 'aligned', 'gathered', 'split', 'itemize', 'enumerate', 'abstract', 'IEEEkeywords')
    blocks = [rf'\\begin\{{{re.escape(name)}\}}.*?\\end\{{{re.escape(name)}\}}' for name in environments]
    blocks += [r'\\includegraphics(?:\[[^]]*\])?\{[^}]+\}', r'\\\[.*?\\\]',
               r'(?<!\\)\$\$.*?(?<!\\)\$\$', r'\\(?:sub)*section\*?\{[^{}]+\}']
    tokens = re.split('(' + '|'.join(blocks) + ')', source, flags=re.S)
    equation_label_map, equation_numbers = equation_labels(tokens)
    asset_label_map, asset_numbers = asset_labels(tokens)
    labels = {key: ('Equation', number) for key, number in equation_label_map.items()}
    labels.update(asset_label_map)
    for token_index, token in enumerate(tokens):
        if not token.strip(): continue
        heading = re.fullmatch(r'\\(section|subsection|subsubsection)\*?\{([^{}]+)\}', token.strip())
        if heading:
            doc.add_heading(clean(heading.group(2)), {'section': 1, 'subsection': 2, 'subsubsection': 3}[heading.group(1)])
            continue
        if token.startswith('\\begin{equation') or token.startswith('\\begin{align') or token.startswith('\\begin{displaymath') or token.startswith('\\begin{gathered') or token.startswith('\\begin{split') or token.startswith(r'\[') or token.startswith('$$'):
            for line_index, line in enumerate(equation_lines(token)):
                paragraph = doc.add_paragraph()
                number = equation_numbers.get((token_index, line_index))
                if number:
                    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.2), WD_TAB_ALIGNMENT.RIGHT)
                math_source = re.sub(r'\\(?:label\{[^}]+\}|nonumber\b|notag\b)', '', line).replace('&', '')
                add_equation(paragraph, math_source)
                if number:
                    paragraph.add_run('\t')
                    bookmark_number(paragraph, number)
                report['equations'] += 1
            continue
        if token.startswith('\\begin{itemize}') or token.startswith('\\begin{enumerate}'):
            style = 'List Bullet' if token.startswith('\\begin{itemize}') else 'List Number'
            for item in re.split(r'\\item\s*', token)[1:]:
                item = re.sub(r'\\end\{(?:itemize|enumerate)\}.*$', '', item, flags=re.S).strip()
                if item: add_text(doc.add_paragraph(style=style), item, entries, numbering, report, labels)
            continue
        if token.startswith('\\begin{abstract}'):
            doc.add_heading('Abstract', 1)
            abstract = re.sub(r'^\\begin\{abstract\}|\\end\{abstract\}$', '', token.strip())
            add_text(doc.add_paragraph(), re.sub(r'\s*\n\s*', ' ', abstract), entries, numbering, report, labels)
            continue
        if token.startswith('\\begin{IEEEkeywords}'):
            keywords = re.sub(r'^\\begin\{IEEEkeywords\}|\\end\{IEEEkeywords\}$', '', token.strip()).strip()
            if keywords:
                paragraph = doc.add_paragraph()
                paragraph.add_run('Keywords: ').bold = True
                add_text(paragraph, keywords, entries, numbering, report, labels)
            continue
        if token.startswith('\\begin{figure'):
            images = re.findall(r'\\includegraphics(?:\[[^]]*\])?\{([^}]+)\}', token)
            caption = re.search(r'\\caption\{([^{}]*)\}', token)
            for image_name in images:
                insert_figure(doc, image_name, files, main, graphic_paths, report)
            if caption:
                paragraph = doc.add_paragraph()
                caption_number(paragraph, *asset_numbers[token_index])
                add_text(paragraph, caption.group(1), entries, numbering, report, labels)
            continue
        if token.startswith('\\includegraphics'):
            image = re.search(r'\{([^}]+)\}', token)
            if image: insert_figure(doc, image.group(1), files, main, graphic_paths, report)
            continue
        if token.startswith('\\begin{algorithm'):
            caption = re.search(r'\\caption\{([^{}]*)\}', token)
            if caption:
                paragraph = doc.add_paragraph()
                caption_number(paragraph, *asset_numbers[token_index])
                add_text(paragraph, caption.group(1), entries, numbering, report, labels)
            body = re.sub(r'\\(?:begin|end)\{(?:algorithm\*?|algorithmic)\}(?:\[[^]]*\])?', '', token)
            body = re.sub(r'\\(?:caption|label)\{[^}]*\}', '', body)
            steps = re.split(r'(?=\\(?:State|Require|Ensure|If|ElsIf|Else|EndIf|For|EndFor|While|EndWhile|Return|Procedure|EndProcedure|KwData|KwResult)\b)', body)
            steps = [re.sub(r'\\;\s*$', '', step.strip()) for step in steps if step.strip()]
            if steps:
                algorithm_table = doc.add_table(rows=len(steps), cols=2)
                algorithm_table.style = 'Table Grid'
                depth = 0
                for index, step in enumerate(steps):
                    match = re.match(r'\\(State|Require|Ensure|If|ElsIf|Else|EndIf|For|EndFor|While|EndWhile|Return|Procedure|EndProcedure|KwData|KwResult)\b', step)
                    command = match.group(1) if match else 'State'
                    content = step[match.end():].strip() if match else step
                    if command in ('If', 'ElsIf', 'For', 'While', 'Procedure') and content.startswith('{') and content.endswith('}'):
                        content = content[1:-1]
                    if command in ('EndIf', 'EndFor', 'EndWhile', 'EndProcedure', 'Else', 'ElsIf'): depth = max(0, depth - 1)
                    algorithm_table.cell(index, 0).text = str(index + 1)
                    paragraph = algorithm_table.cell(index, 1).paragraphs[0]
                    paragraph.paragraph_format.left_indent = Inches(min(depth, 5) * .18)
                    if command != 'State':
                        lead = paragraph.add_run(command.replace('Kw', '') + ' ')
                        lead.bold = True
                    add_text(paragraph, content, entries, numbering, report, labels)
                    if command in ('If', 'For', 'While', 'Procedure', 'Else', 'ElsIf'): depth += 1
                report['algorithms'] += 1
            continue
        if token.startswith('\\begin{table') or token.startswith('\\begin{tabular') or token.startswith('\\begin{longtable'):
            caption = re.search(r'\\caption\{([^{}]*)\}', token)
            if caption:
                paragraph = doc.add_paragraph()
                caption_number(paragraph, *asset_numbers[token_index])
                add_text(paragraph, caption.group(1), entries, numbering, report, labels)
            body = tabular_body(token)
            if body is not None:
                body = protect_shortstacks(body)
                rows = [r.strip() for r in re.split(r'(?<!\\)\\\\', body) if r.strip()]
                parsed = [[c.strip() for c in re.split(r'(?<!\\)&', re.sub(r'\\(?:hline|toprule|midrule|bottomrule)', '', row))] for row in rows]
                parsed = [r for r in parsed if any(r)]
                if parsed:
                    table = doc.add_table(rows=len(parsed), cols=max(map(len, parsed)))
                    table.style = 'Table Grid'
                    for i, row in enumerate(parsed):
                        for j, value in enumerate(row):
                            table_cell_content(table.cell(i, j).paragraphs[0], value, entries, numbering, report, labels)
                    report['tables'] += 1
            continue
        token = re.sub(r'\\(?:maketitle|bibliography|bibliographystyle|usepackage|documentclass)(?:\[[^]]*\])?\{[^}]*\}', '', token)
        for paragraph in re.split(r'\n\s*\n', token):
            paragraph = re.sub(r'\s*\n\s*', ' ', paragraph).strip()
            if paragraph and not re.fullmatch(r'(?:[pmb]\{[^{}]*\}|[clr|{}\\\s])+', paragraph):
                add_text(doc.add_paragraph(), paragraph, entries, numbering, report, labels)
    if numbering:
        doc.add_heading('References', 1)
        references = doc.add_paragraph()
        field(references, ' ADDIN EN.REFLIST ', '\n'.join(f'[{n}] {entries[k].get("raw") or (entries[k].get("author", k) + " (" + entries[k].get("year", "") + "). " + entries[k].get("title", k) + ".")}' for k, n in numbering.items()))
    if not entries: report['warnings'].append('No .bib references found; citations could not become EndNote fields.')
    if report['missing_figures']: report['warnings'].append('Some figures could not be inserted.')
    if re.search(r'\\(?:begin\{(?:tikzpicture|longtable)|newcommand)', source):
        report['warnings'].append('TikZ, longtable, or custom macros may need manual cleanup.')
    known = {'begin','end','section','subsection','subsubsection','cite','citep','citet','citeauthor','citeyear','ref','eqref','autoref','cref','Cref','includegraphics','caption','label','index','emph','textit','textbf','underline','texttt','shortstack','alpha','beta','gamma','delta','epsilon','theta','lambda','mu','pi','sigma','phi','omega','Gamma','Delta','Sigma','Omega','times','cdot','pm','leq','geq','neq','infty','partial','sum','int','rightarrow','in','approx','frac','dfrac','tfrac','binom','sqrt','overline','bar','hat','vec','tilde','overbrace','underbrace','mathbb','mathcal','mathfrak','sin','cos','tan','log','ln','exp','lim','max','min','mathrm','mathbf','mathit','text','operatorname','boldsymbol','left','right','maketitle','bibliography','bibliographystyle','usepackage','documentclass','title','author','item','hline','toprule','midrule','bottomrule'}
    known.update(GREEK)
    known.update(SYMBOLS)
    known.update({'State','Require','Ensure','If','ElsIf','Else','EndIf','For','EndFor','While','EndWhile','Return','Procedure','EndProcedure','KwData','KwResult',
                  'quad','qquad','big','Big','bigg','Bigg','bigl','bigr','Bigl','Bigr','biggl','biggr','Biggl','Biggr',
                  'widetilde','tanh'})
    unknown = sorted({name for name in re.findall(r'\\([A-Za-z]+)', source) if name not in known})
    if unknown: report['warnings'].append('Review unsupported LaTeX commands: ' + ', '.join(unknown[:20]))
    report['unresolved_citations'] = sorted(set(report['unresolved_citations']))
    report['unresolved_references'] = sorted(set(report['unresolved_references']))
    return doc, report
