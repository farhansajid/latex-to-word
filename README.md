# LaTeX to Word

A local web interface for converting a LaTeX manuscript, its `.bib` file, and figure assets into a Word document. Citations are written as EndNote `ADDIN EN.CITE` fields with embedded reference records; a bibliography field is also included. Open the result in Microsoft Word with EndNote Cite While You Write and use **Update Citations and Bibliography** to verify and refresh the fields.

## Run

```powershell
python app.py
```

Open <http://localhost:8765>. Install dependencies with `pip install -r requirements.txt`. No files are uploaded to a remote service.

The app inspects the selected project before conversion, shows the `.tex`, `.bib`, and figure inventory, and lets you choose the main `.tex` file when several are present. The result panel shows counts and any missing assets or unresolved references. ZIP uploads are limited to 40 MB compressed, 100 MB unpacked, and 500 entries.

Upload a `.tex` file or a `.zip` containing the main `.tex`, `.bib`, and images. The converter handles headings, paragraphs, bold/italic/underlined/monospace text, lists, abstracts, common inline and display equations as editable Word math, citations from `.bib` or inline `\bibitem` entries, figures (including first-page PDF rendering), and `tabular` tables. Its report lists unresolved citations, missing images, and unsupported commands. The Word math translator covers common symbols, fractions, roots, and scripts; complex LaTeX packages and macros may need manual cleanup in Word. EndNote compatibility requires a final check in Word with the EndNote add-in; this environment cannot run that add-in.

Numbered `equation` and `align` lines receive Word `SEQ Equation` fields beside the math. Figures and tables receive numbered `SEQ Figure` and `SEQ Table` captions. `\label{...}` becomes a bookmark on its number; `\ref{...}`, `\eqref{...}`, `\autoref{...}`, and `\cref{...}` become Word `REF` fields, including forward references. In Word, select all and press **F9** after editing equations, figures, or tables to refresh numbers and references.

Math conversion recognizes fractions (`\frac`, `\dfrac`, `\tfrac`), binomial coefficients, square and indexed roots, combined superscripts and subscripts, accents (`\hat`, `\bar`, `\vec`, `\tilde`, `\overline`), scalable `\left`/`\right` delimiters, matrix environments (`matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `cases`), common symbols, and text/operator commands. Unsupported math macros remain visible for manual review.

Inline `$...$` and display `$$...$$` become editable Word math; an escaped `\$` stays a currency symbol. Table captions appear above their tables. PNG and JPEG images are embedded directly, while PDF figures are rendered to PNG before insertion.

Figures are resolved relative to the main `.tex` file and directories declared with `\graphicspath`, including extensionless `\includegraphics` calls. Standalone `\includegraphics` commands are inserted as well. Put image files in the ZIP project, or choose them in the figure picker when converting a single `.tex` file. Missing files are listed in the conversion report.

Table cells use the same inline converter as ordinary paragraphs, including `$...$` math, text formatting, citations, and cross-references. Common `algorithm`/`algorithmic` commands become a numbered, indented Word table with bold step keywords and editable math; algorithm captions and `\label`/`\ref` use Word fields.

Nested text styles such as `\textbf{bold \textit{and italic}}` become Word run formatting. In tables, `\shortstack[l]{First\\Second}` stays in one cell with a Word line break and the requested alignment. Grouped math such as `{a+b}^2` keeps the exponent on the full group, and `\mathbf`, `\boldsymbol`, and `\textbf` inside math produce bold Word math runs.

Document setup commands such as `\title` and `\maketitle` are consumed without leaking raw LaTeX into the DOCX. Empty `IEEEkeywords` blocks, comments, labels, and isolated table column specifications are removed. `tabularx` and `longtable` column declarations are parsed, and `aligned` math becomes a Word equation array. Math also supports `\mathbb`, `\mathcal`, `\mathfrak`, braces, and common function names.
