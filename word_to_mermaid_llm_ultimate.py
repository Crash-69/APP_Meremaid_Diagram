#!/usr/bin/env python3
"""App locale per convertire DOCX o Markdown in diagrammi Mermaid."""

from __future__ import annotations

import base64
import html
import io
import json
import mimetypes
import re
import zipfile
from pathlib import PurePath
from types import SimpleNamespace
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from xml.etree import ElementTree

import streamlit as st

try:
    from markitdown import MarkItDown, StreamInfo
except ImportError:
    MarkItDown = None
    StreamInfo = None

try:
    from magika import Magika
except ImportError:
    Magika = None


# Configurazione locale, limiti di sicurezza e namespace OOXML.
OLLAMA_URL = "http://localhost:11434"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAX_ARCHIVE_ENTRIES = 5000
MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class OllamaCompletions:
    """Adattatore leggero OpenAI-compatibile per il descrittore immagini MarkItDown."""

    def __init__(self) -> None:
        self.chat = SimpleNamespace(completions=self)

    def create(self, *, model: str, messages: list[dict], **options: object) -> SimpleNamespace:
        """Invia a Ollama una richiesta multimodale e presenta la risposta attesa."""
        payload = {"model": model, "messages": messages, "stream": False}
        payload.update(options)
        request = Request(
            f"{OLLAMA_URL}/v1/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
        content = result["choices"][0]["message"].get("content", "")
        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))]
        )


@st.cache_data(ttl=20, show_spinner=False)
def get_ollama_models() -> tuple[list[str], str | None]:
    """Interroga Ollama per l'elenco modelli; un errore non blocca il parser locale."""
    request = Request(f"{OLLAMA_URL}/api/tags", headers={"Accept": "application/json"})
    try:
        with urlopen(request, timeout=3) as response:
            payload = json.loads(response.read().decode("utf-8"))
        models = sorted(
            {str(item["name"]) for item in payload.get("models", []) if item.get("name")}
        )
        return models, None
    except (HTTPError, URLError, TimeoutError, OSError, ValueError, KeyError) as exc:
        return [], str(exc)


def _word_tag(local_name: str) -> str:
    """Restituisce il tag OOXML con namespace WordprocessingML."""
    return f"{{{WORD_NS}}}{local_name}"


def validate_docx(content: bytes) -> None:
    """Valida un archivio DOCX in memoria e limita i rischi di decompressione."""
    if not content:
        raise ValueError("Il file caricato è vuoto.")
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("Il file supera il limite consentito di 20 MB.")
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_ARCHIVE_ENTRIES:
                raise ValueError("Il DOCX contiene un numero eccessivo di elementi.")
            if sum(entry.file_size for entry in entries) > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("Il DOCX supera il limite di decompressione consentito.")
            required = {"[Content_Types].xml", "word/document.xml"}
            if not required.issubset(archive.namelist()):
                raise ValueError("Il file non contiene la struttura di un documento DOCX.")
    except zipfile.BadZipFile as exc:
        raise ValueError("Il file Word è corrotto o non è un archivio DOCX valido.") from exc


def _read_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    """Legge gli hyperlink esterni contenuti nel documento Word."""
    try:
        root = ElementTree.fromstring(archive.read("word/_rels/document.xml.rels"))
    except (KeyError, ElementTree.ParseError):
        return {}
    return {
        relation.get("Id", ""): relation.get("Target", "")
        for relation in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
        if relation.get("Id")
        and relation.get("Target")
        and relation.get("TargetMode") == "External"
    }


def _paragraph_markdown(
    paragraph: ElementTree.Element, relationships: dict[str, str]
) -> tuple[str, int, bool]:
    """Estrae testo, formattazione essenziale, titoli, liste e alt-text immagini."""
    fragments: list[str] = []
    for child in paragraph:
        if child.tag == _word_tag("r"):
            text = "".join(node.text or "" for node in child.iter(_word_tag("t")))
            if text:
                properties = child.find(_word_tag("rPr"))
                if properties is not None:
                    if properties.find(_word_tag("b")) is not None:
                        text = f"**{text}**"
                    if properties.find(_word_tag("i")) is not None:
                        text = f"*{text}*"
                    if properties.find(_word_tag("strike")) is not None:
                        text = f"~~{text}~~"
                fragments.append(text)
            if child.find(_word_tag("tab")) is not None:
                fragments.append("    ")
            if child.find(_word_tag("br")) is not None:
                fragments.append("  \n")
        elif child.tag == _word_tag("hyperlink"):
            text = "".join(node.text or "" for node in child.iter(_word_tag("t")))
            target = relationships.get(child.get(f"{{{REL_NS}}}id", ""))
            fragments.append(f"[{text}]({target})" if target else text)
        elif child.tag in (_word_tag("drawing"), _word_tag("pict")):
            description = next(
                (
                    node.get("descr") or node.get("title")
                    for node in child.iter()
                    if node.get("descr") or node.get("title")
                ),
                "immagine",
            )
            fragments.append(f"[Immagine: {description}]")

    text = "".join(fragments).strip()
    if not text:
        return "", 0, False
    properties = paragraph.find(_word_tag("pPr"))
    style_name = ""
    list_level = 0
    is_list_item = False
    if properties is not None:
        style = properties.find(_word_tag("pStyle"))
        if style is not None:
            style_name = style.get(_word_tag("val"), "")
        numbering = properties.find(_word_tag("numPr"))
        if numbering is not None:
            is_list_item = numbering.find(_word_tag("numId")) is not None
            level = numbering.find(_word_tag("ilvl"))
            if level is not None:
                list_level = int(level.get(_word_tag("val"), "0"))
    heading = re.search(r"heading\s*([1-6])", style_name, re.IGNORECASE)
    if heading:
        return f"{'#' * int(heading.group(1))} {text}", 0, False
    if style_name.lower() in {"title", "titolo"}:
        return f"# {text}", 0, False
    if is_list_item:
        return f"{'  ' * list_level}- {text}", list_level, True
    return text, 0, False


def convert_docx_fallback(content: bytes) -> str:
    """Converte contenuti OOXML essenziali senza creare file temporanei."""
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        try:
            document = ElementTree.fromstring(archive.read("word/document.xml"))
        except (KeyError, ElementTree.ParseError) as exc:
            raise ValueError("Il documento Word non contiene un corpo leggibile.") from exc
        body = document.find(_word_tag("body"))
        if body is None:
            raise ValueError("Il documento Word non contiene un corpo leggibile.")
        relationships = _read_relationships(archive)
        lines: list[str] = []
        for block in body:
            if block.tag == _word_tag("p"):
                text, _, is_list_item = _paragraph_markdown(block, relationships)
                if text:
                    lines.append(text if is_list_item else f"\n{text}\n")
            elif block.tag == _word_tag("tbl"):
                table_rows: list[list[str]] = []
                for row in block.findall(_word_tag("tr")):
                    cells = [
                        " ".join(
                            "".join(node.text or "" for node in paragraph.iter(_word_tag("t")))
                            for paragraph in cell.findall(_word_tag("p"))
                        ).strip().replace("|", "\\|")
                        for cell in row.findall(_word_tag("tc"))
                    ]
                    if cells:
                        table_rows.append(cells)
                if table_rows:
                    width = max(map(len, table_rows))
                    lines.append("| " + " | ".join(table_rows[0]) + " |")
                    lines.append("| " + " | ".join(["---"] * width) + " |")
                    lines.extend("| " + " | ".join(row) + " |" for row in table_rows[1:])
        markdown = "\n".join(lines).strip()
    if not markdown:
        raise ValueError("Il documento Word non contiene testo convertibile.")
    return markdown + "\n"


def describe_docx_images(content: bytes, model: str) -> tuple[list[str], list[str]]:
    """Descrive immagini incorporate inviandone i byte a Ollama senza file temporanei."""
    supported_types = {"image/png", "image/jpeg", "image/gif", "image/webp"}
    descriptions: list[str] = []
    warnings: list[str] = []
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        image_names = sorted(
            name for name in archive.namelist() if name.startswith("word/media/")
        )
        if len(image_names) > 20:
            warnings.append("Sono state considerate al massimo 20 immagini.")
        for index, name in enumerate(image_names[:20], start=1):
            mime_type, _ = mimetypes.guess_type(name)
            if mime_type not in supported_types:
                warnings.append(f"Formato immagine non supportato: {PurePath(name).suffix or name}.")
                continue
            image_bytes = archive.read(name)
            if not image_bytes or len(image_bytes) > 8 * 1024 * 1024:
                warnings.append(f"Immagine {index} vuota o superiore a 8 MB: ignorata.")
                continue
            data_uri = f"data:{mime_type};base64,{base64.b64encode(image_bytes).decode('ascii')}"
            messages = [{
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Descrivi in italiano questa immagine incorporata in un documento. "
                            "Trascrivi il testo leggibile, poi riassumi gli elementi visivi. "
                            "Non inventare dettagli illeggibili."
                        ),
                    },
                    {"type": "image_url", "image_url": {"url": data_uri}},
                ],
            }]
            try:
                response = OllamaCompletions().create(model=model, messages=messages)
                description = response.choices[0].message.content.strip()
                if description:
                    descriptions.append(f"**Immagine {index}:** {description}")
            except (HTTPError, URLError, TimeoutError, OSError, ValueError, KeyError) as exc:
                warnings.append(f"Descrizione immagine {index} non riuscita: {exc}")
    return descriptions, warnings


def _append_image_descriptions(
    markdown: str, content: bytes, model: str | None
) -> tuple[str, str | None]:
    """Aggiunge descrizioni LLM al Markdown convertito quando è selezionato un modello."""
    if not model:
        return markdown, None
    descriptions, warnings = describe_docx_images(content, model)
    if descriptions:
        markdown += "\n## Immagini\n\n" + "\n\n".join(descriptions) + "\n"
    status = f"Immagini descritte da Ollama: {len(descriptions)}."
    if warnings:
        status += " " + " ".join(warnings)
    return markdown, status if descriptions or warnings else None


def _fallback_with_image_descriptions(
    content: bytes, model: str | None
) -> tuple[str, str | None]:
    """Converte OOXML localmente e poi arricchisce il risultato con immagini LLM."""
    markdown = convert_docx_fallback(content)
    return _append_image_descriptions(markdown, content, model)


def convert_docx(content: bytes, filename: str, model: str | None) -> tuple[str, str]:
    """Usa MarkItDown su un flusso in memoria, con fallback OOXML locale."""
    validate_docx(content)
    if Magika is not None:
        detected = Magika().identify_bytes(content).output
        if detected.label != "docx" or detected.mime_type != DOCX_MIME:
            raise ValueError(
                f"Il file è identificato come {detected.label} "
                f"({detected.mime_type}), non come DOCX."
            )
    if MarkItDown is not None:
        try:
            options: dict[str, object] = {}
            if model:
                options.update(
                    llm_client=OllamaCompletions(),
                    llm_model=model,
                    llm_prompt=(
                        "Descrivi in italiano gli elementi visivi rilevanti. "
                        "Trascrivi il testo leggibile presente nelle immagini; "
                        "se non è leggibile, dichiaralo senza inventare contenuti."
                    ),
                )
            info = StreamInfo(
                mimetype=DOCX_MIME,
                extension=".docx",
                filename=PurePath(filename).name,
            )
            result = MarkItDown().convert(io.BytesIO(content), stream_info=info, **options)
            markdown = result.text_content.strip()
            if markdown:
                markdown, image_status = _append_image_descriptions(
                    markdown + "\n", content, model
                )
                engine = "MarkItDown"
                if image_status:
                    engine += f"; {image_status}"
                return markdown, engine
        except Exception as exc:
            message = str(exc).lower()
            missing_docx_extra = (
                "missingdependencyexception" in type(exc).__name__.lower()
                or "dependencies needed to read .docx" in message
                or "optional dependency [docx]" in message
            )
            fallback, image_status = _fallback_with_image_descriptions(content, model)
            if missing_docx_extra:
                engine = "Fallback OOXML locale (extra DOCX di MarkItDown non installato)"
            else:
                engine = f"Fallback OOXML locale (errore MarkItDown: {type(exc).__name__})"
            if image_status:
                engine += f"; {image_status}"
            return fallback, engine
    fallback, image_status = _fallback_with_image_descriptions(content, model)
    engine = "Fallback OOXML locale"
    if image_status:
        engine += f"; {image_status}"
    return fallback, engine


def convert_markdown(content: bytes) -> str:
    """Decodifica file Markdown UTF-8 direttamente dalla memoria."""
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("Il file supera il limite consentito di 20 MB.")
    try:
        markdown = content.decode("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
    except UnicodeDecodeError as exc:
        raise ValueError("Il file Markdown deve essere codificato in UTF-8.") from exc
    if not markdown.strip():
        raise ValueError("Il file Markdown è vuoto.")
    return markdown if markdown.endswith("\n") else markdown + "\n"


def _mermaid_label(value: str) -> str:
    """Pulisce caratteri di controllo e codifica le etichette per Mermaid."""
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", value)
    value = re.sub(r"\s+", " ", value).strip() or "(elemento senza testo)"
    return html.escape(value, quote=True)


def markdown_to_mermaid(markdown: str) -> str:
    """Costruisce graph TD collegando titoli e liste anche annidate."""
    rows = ["graph TD", '  n0["Documento"]']
    headings: list[tuple[int, str]] = []
    list_nodes: list[tuple[int, str]] = []
    next_id = 1
    for line in markdown.splitlines():
        heading = re.match(r"^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$", line)
        if heading:
            level = len(heading.group(1))
            while headings and headings[-1][0] >= level:
                headings.pop()
            parent = headings[-1][1] if headings else "n0"
            node = f"n{next_id}"
            next_id += 1
            rows.extend([f'  {node}["{_mermaid_label(heading.group(2))}"]', f"  {parent} --> {node}"])
            headings.append((level, node))
            list_nodes.clear()
            continue
        item = re.match(r"^([ \t]*)(?:[-+*]|\d+[.)])\s+(.+?)\s*$", line)
        if not item:
            continue
        indent = len(item.group(1).expandtabs(4))
        while list_nodes and list_nodes[-1][0] >= indent:
            list_nodes.pop()
        parent = (
            list_nodes[-1][1]
            if list_nodes
            else headings[-1][1]
            if headings
            else "n0"
        )
        node = f"n{next_id}"
        next_id += 1
        rows.extend([f'  {node}["{_mermaid_label(item.group(2))}"]', f"  {parent} --> {node}"])
        list_nodes.append((indent, node))
    return "\n".join(rows)


def _valid_mermaid(diagram: str) -> bool:
    """Verifica che la risposta LLM contenga solo la grammatica generata qui."""
    rows = [line.strip() for line in diagram.splitlines() if line.strip()]
    if not rows or rows[0] not in {"graph TD", "flowchart TD"}:
        return False
    declared: set[str] = set()
    edges: list[str] = []
    for row in rows[1:]:
        node = re.fullmatch(r'(n\d+)\["[^\n]*"\]', row)
        edge = re.fullmatch(r"(n\d+)\s+-->\s+(n\d+)", row)
        if node:
            declared.add(node.group(1))
        elif edge:
            edges.extend(edge.groups())
        else:
            return False
    return all(endpoint in declared for endpoint in edges)


def validate_mermaid_with_ollama(diagram: str, model: str) -> tuple[str, str | None]:
    """Chiede una correzione sintattica e scarta risposte LLM non conformi."""
    system = (
        'Correggi solo la sintassi Mermaid. Restituisci esclusivamente il diagramma, '
        'senza Markdown o spiegazioni. Mantieni graph TD, ID nN, nodi nN["etichetta"] '
        'e archi nN --> nN.'
    )
    payload = {
        "model": model,
        "stream": False,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": diagram[:50000]},
        ],
        "options": {"temperature": 0, "num_predict": 3000},
    }
    request = Request(
        f"{OLLAMA_URL}/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=90) as response:
            candidate = json.loads(response.read().decode("utf-8"))["message"]["content"].strip()
        candidate = re.sub(r"^```(?:mermaid)?\s*|\s*```$", "", candidate, flags=re.I)
        if _valid_mermaid(candidate):
            return candidate, None
        return diagram, "La risposta LLM non ha superato i controlli: uso il diagramma locale."
    except (HTTPError, URLError, TimeoutError, OSError, ValueError, KeyError) as exc:
        return diagram, f"Validazione Ollama non disponibile: {exc}"


def make_standalone_html(diagram: str, title: str) -> str:
    """Genera HTML interattivo con Mermaid caricato dal CDN jsDelivr."""
    safe_title = html.escape(title, quote=True)
    safe_diagram = html.escape(diagram, quote=False)
    return f'''<!doctype html>
<html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>{safe_title}</title>
<style>body{{margin:0;padding:2rem;background:#f5f5f5;color:#0f172a;font-family:system-ui,sans-serif}}main{{max-width:1100px;margin:auto;padding:1.5rem;background:#fff;border:1px solid #e2e8f0;border-radius:12px;overflow:auto}}h1{{color:#047857;font-size:1.4rem}}</style>
<script type="module">
import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs';
mermaid.initialize({{startOnLoad:true,securityLevel:'strict'}});
</script></head><body><main><h1>{safe_title}</h1>
<pre class="mermaid">{safe_diagram}</pre></main></body></html>
'''


def _safe_stem(filename: str) -> str:
    """Sanifica il nome del documento per i download."""
    stem = PurePath(filename.replace("\\", "/")).stem
    return re.sub(r"[^\w.-]+", "_", stem, flags=re.UNICODE).strip("._") or "documento"


def main() -> None:
    """Disegna l'interfaccia Streamlit e gestisce entrambi i formati in ingresso."""
    st.set_page_config(page_title="Word → Mermaid", page_icon="🌿", layout="wide")
    st.markdown(
        """<style>
        :root{color-scheme:light;--text-color:#0f172a;--background-color:#f5f5f5;--secondary-background-color:#fff;--primary-color:#059669}
        [data-testid="stAppViewContainer"],.stApp{background:#f5f5f5!important;color:#0f172a!important}
        [data-testid="stHeader"]{background:transparent}
        [data-testid="stSidebar"]{background:#fff!important;border-right:1px solid #e2e8f0}
        [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"]{color:#334155!important}
        [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3{color:#0f172a!important}
        .block-container{max-width:1120px;padding:2.25rem 2rem 3rem}
        h1,h2,h3{color:#0f172a!important}
        h1{font-size:1.75rem!important;line-height:1.2}
        p,label,[data-testid="stWidgetLabel"],[data-testid="stCaptionContainer"],.stMarkdown{color:#334155}
        [data-testid="stCaptionContainer"] p{color:#64748b!important}
        div[data-baseweb="select"]>div{background:#fff!important;border-color:#cbd5e1!important;color:#0f172a!important}
        div[data-baseweb="select"] span{color:#0f172a!important}
        [data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:2px dashed #cbd5e1!important;border-radius:12px}
        [data-testid="stFileUploaderDropzone"] p,[data-testid="stFileUploaderDropzone"] span{color:#334155!important}
        [data-testid="stFileUploaderDropzone"] button{background:#0f172a!important;border:0!important;border-radius:9px;color:#fff!important}
        [data-testid="stFileUploaderDropzone"] button *{color:#fff!important}
        div.stButton>button,div[data-testid="stDownloadButton"] button{border-radius:10px;font-weight:600}
        div.stButton>button[kind="primary"]{background:#059669!important;border-color:#059669!important;color:#fff!important}
        div.stButton>button[kind="primary"]:hover{background:#047857!important;border-color:#047857!important}
        div.stButton>button:disabled{background:#cbd5e1!important;color:#475569!important}
        div[data-testid="stDownloadButton"] button{background:#0f172a;color:#fff}
        div[data-testid="stDownloadButton"] button:hover{background:#1e293b;color:#fff}
        [data-testid="stAlert"]{border-radius:12px}
        [data-testid="stAlert"] [data-testid="stMarkdownContainer"],
        [data-testid="stAlert"] [data-testid="stMarkdownContainer"] p{color:#334155!important}
        [data-testid="stTabs"] button[data-baseweb="tab"]{color:#475569}
        [data-testid="stTabs"] button[aria-selected="true"]{color:#047857}
        [data-testid="stTabs"] [data-baseweb="tab-highlight"]{background:#059669}
        @media(max-width:700px){.block-container{padding:1.25rem 1rem 2rem}}
        </style>""",
        unsafe_allow_html=True,
    )
    st.title("Word in Markdown · Mermaid")
    st.caption("Conversione locale di documenti Word o Markdown in diagrammi gerarchici.")

    models, ollama_error = get_ollama_models()
    with st.sidebar:
        st.subheader("Modello locale")
        if models:
            model = st.selectbox("Modello Ollama", models, key="ollama_model")
            st.caption("Ollama è raggiungibile in locale.")
        else:
            model = None
            st.selectbox("Modello Ollama", ["Solo parser locale"], disabled=True)
            st.warning("Ollama non è raggiungibile. Conversione e parser locale restano disponibili.")
            if ollama_error:
                st.caption(ollama_error)
        st.divider()
        st.caption("Elaborazione in memoria; nessun file temporaneo.")

    uploaded = st.file_uploader(
        "Documento Word o Markdown",
        type=["docx", "md"],
        max_upload_size=20,
        help="Formati accettati: .docx e .md · massimo 20 MB.",
    )
    if st.button("Analizza documento", type="primary", disabled=uploaded is None):
        try:
            with st.spinner("Elaborazione in corso…"):
                content = uploaded.getvalue()
                if uploaded.name.lower().endswith(".md"):
                    markdown = convert_markdown(content)
                    engine = "Markdown caricato"
                else:
                    markdown, engine = convert_docx(content, uploaded.name, model)
                diagram = markdown_to_mermaid(markdown)
                llm_warning = None
                if model:
                    diagram, llm_warning = validate_mermaid_with_ollama(diagram, model)
            st.session_state["conversion_result"] = {
                "markdown": markdown,
                "diagram": diagram,
                "engine": engine,
                "filename": uploaded.name,
                "llm_warning": llm_warning,
            }
        except Exception as exc:
            st.error(f"Elaborazione non riuscita: {exc}")

    result = st.session_state.get("conversion_result")
    if not result:
        st.info("Carica un documento .docx o .md per iniziare.")
        return
    if result["engine"].startswith("Fallback"):
        st.warning(f"Conversione eseguita con il parser OOXML locale. {result['engine']}")
    else:
        st.success(f"Documento pronto: {result['engine']}.")
    if result.get("llm_warning"):
        st.warning(result["llm_warning"])

    markdown_tab, diagram_tab = st.tabs(["Testo Markdown", "Diagramma"])
    with markdown_tab:
        st.code(result["markdown"], language="markdown", line_numbers=True)
    with diagram_tab:
        st.mermaid_chart(result["diagram"])
        with st.expander("Codice Mermaid"):
            st.code(result["diagram"], language="mermaid")

    stem = _safe_stem(result["filename"])
    st.subheader("Download")
    columns = st.columns(3)
    columns[0].download_button(
        "Scarica Markdown", result["markdown"].encode("utf-8"),
        f"{stem}.md", "text/markdown", use_container_width=True,
    )
    columns[1].download_button(
        "Scarica Mermaid", result["diagram"].encode("utf-8"),
        f"{stem}.mmd", "text/plain", use_container_width=True,
    )
    columns[2].download_button(
        "Scarica HTML interattivo", make_standalone_html(result["diagram"], stem).encode("utf-8"),
        f"{stem}_diagramma.html", "text/html", use_container_width=True,
    )


if __name__ == "__main__":
    main()