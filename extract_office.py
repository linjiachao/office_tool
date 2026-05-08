# -*- coding: utf-8 -*-
from pathlib import Path
from pathlib import PurePosixPath
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import csv
import posixpath
import re
import sys
import traceback
from dataclasses import dataclass


BASE = Path(__file__).resolve().parent
DEFAULT_OUT_DIR = BASE / "converted"
SUPPORTED_SUFFIXES = {".xlsx", ".docx"}

NS_XLSX = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
NS_DOCX = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


@dataclass
class ConversionTask:
    source_path: Path
    output_dir: Path


def safe_filename(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip() or "Sheet"


def col_index(cell_ref):
    match = re.match(r"[A-Z]+", cell_ref or "")
    if not match:
        return 0

    value = 0
    for ch in match.group(0):
        value = value * 26 + ord(ch) - ord("A") + 1
    return value - 1


def normalize_sheet_path(target):
    target = (target or "").replace("\\", "/").lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target


def cell_text(cell, shared_strings):
    cell_type = cell.get("t")

    if cell_type == "s":
        value = cell.find("main:v", NS_XLSX)
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return value.text

    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//main:t", NS_XLSX))

    if cell_type == "b":
        value = cell.find("main:v", NS_XLSX)
        return "TRUE" if value is not None and value.text == "1" else "FALSE"

    value = cell.find("main:v", NS_XLSX)
    return value.text if value is not None and value.text is not None else ""


def read_shared_strings(zip_file):
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("main:si", NS_XLSX):
        strings.append("".join(text.text or "" for text in item.findall(".//main:t", NS_XLSX)))
    return strings


def read_xlsx_sheets(path):
    sheets = []

    with ZipFile(path) as zip_file:
        shared_strings = read_shared_strings(zip_file)

        workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
        rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
        relmap = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.findall(f"{REL_NS}Relationship")
        }

        for sheet in workbook.findall(".//main:sheet", NS_XLSX):
            sheet_name = sheet.get("name", "Sheet")
            rel_id = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheet_target = relmap.get(rel_id)
            if not sheet_target:
                continue

            sheet_path = normalize_sheet_path(sheet_target)
            if sheet_path not in zip_file.namelist():
                sheets.append((sheet_name, [[f"Skipped: cannot find {sheet_path}"]]))
                continue

            root = ET.fromstring(zip_file.read(sheet_path))
            rows = []

            for row in root.findall(".//main:sheetData/main:row", NS_XLSX):
                values = []
                for cell in row.findall("main:c", NS_XLSX):
                    index = col_index(cell.get("r"))
                    while len(values) < index:
                        values.append("")
                    values.append(cell_text(cell, shared_strings))

                if any(value != "" for value in values):
                    rows.append(values)

            sheets.append((sheet_name, rows))

    return sheets


def convert_xlsx(path, output_dir):
    written = []
    sheets = read_xlsx_sheets(path)
    output_dir.mkdir(parents=True, exist_ok=True)

    for sheet_name, rows in sheets:
        output_name = f"{path.stem}__{safe_filename(sheet_name)}.csv"
        output_path = output_dir / output_name

        with output_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file)
            writer.writerows(rows)

        written.append(output_path)

    return written


def word_val(element):
    if element is None:
        return None
    return element.get(f"{W}val")


def markdown_escape(text):
    return text.replace("\\", "\\\\").replace("|", "\\|")


def read_docx_relationships(zip_file):
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in zip_file.namelist():
        return {}

    root = ET.fromstring(zip_file.read(rels_path))
    rels = {}
    for rel in root.findall(f"{REL_NS}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        rel_type = rel.get("Type", "")
        if rel_id and target and rel_type.endswith("/image"):
            rels[rel_id] = target
    return rels


def normalize_docx_target(target):
    target = (target or "").replace("\\", "/")
    if target.startswith(("http://", "https://")):
        return target
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("word", target))


class DocxImageExporter:
    def __init__(self, zip_file, rels, media_dir):
        self.zip_file = zip_file
        self.rels = rels
        self.media_dir = media_dir
        self.exported = {}
        self.used_names = set()

    def unique_name(self, target):
        original = safe_filename(PurePosixPath(target.replace("\\", "/")).name)
        if not original:
            original = f"image_{len(self.used_names) + 1:03d}.bin"

        stem = Path(original).stem
        suffix = Path(original).suffix
        candidate = original
        counter = 2
        while candidate.lower() in self.used_names:
            candidate = f"{stem}_{counter}{suffix}"
            counter += 1

        self.used_names.add(candidate.lower())
        return candidate

    def markdown_image(self, rel_id):
        target = self.rels.get(rel_id)
        if not target:
            return ""

        if target.startswith(("http://", "https://")):
            return f"![{rel_id}]({target})"

        if rel_id not in self.exported:
            zip_path = normalize_docx_target(target)
            if zip_path not in self.zip_file.namelist():
                return f"![missing image: {rel_id}]({target})"

            self.media_dir.mkdir(exist_ok=True)
            filename = self.unique_name(target)
            output_path = self.media_dir / filename
            output_path.write_bytes(self.zip_file.read(zip_path))
            self.exported[rel_id] = filename

        filename = self.exported[rel_id]
        alt = Path(filename).stem
        rel_path = f"{self.media_dir.name}/{filename}".replace("\\", "/")
        return f"![{alt}]({rel_path})"


def image_rel_ids(element):
    rel_ids = []
    for node in element.iter():
        for attr, value in node.attrib.items():
            if attr in (f"{R}embed", f"{R}link", f"{R}id") and value:
                if value not in rel_ids:
                    rel_ids.append(value)
    return rel_ids


def paragraph_content_items(paragraph, image_exporter=None):
    items = []
    buffer = []

    def flush_buffer():
        text = "".join(buffer).strip()
        if text:
            items.append(("text", text))
        buffer.clear()

    for node in paragraph.iter():
        if node.tag == f"{W}t":
            buffer.append(node.text or "")
        elif node.tag == f"{W}tab":
            buffer.append("\t")
        elif node.tag in (f"{W}br", f"{W}cr"):
            buffer.append("\n")
        elif node.tag in (f"{W}drawing", f"{W}pict") and image_exporter is not None:
            flush_buffer()
            for rel_id in image_rel_ids(node):
                image = image_exporter.markdown_image(rel_id)
                if image:
                    items.append(("image", image))

    flush_buffer()
    return items


def paragraph_to_markdown_lines(paragraph, image_exporter=None):
    lines = []
    for item_type, value in paragraph_content_items(paragraph, image_exporter):
        if item_type == "text":
            lines.append(value)
        elif item_type == "image":
            lines.append(value)
    return lines


def paragraph_text(paragraph):
    return "".join(
        value for item_type, value in paragraph_content_items(paragraph) if item_type == "text"
    ).strip()


def paragraph_style_id(paragraph):
    style = paragraph.find("./w:pPr/w:pStyle", NS_DOCX)
    return word_val(style)


def paragraph_outline_level(paragraph):
    outline = paragraph.find("./w:pPr/w:outlineLvl", NS_DOCX)
    value = word_val(outline)
    if value is None:
        return None
    try:
        return int(value) + 1
    except ValueError:
        return None


def paragraph_numbering(paragraph):
    num_pr = paragraph.find("./w:pPr/w:numPr", NS_DOCX)
    if num_pr is None:
        return None

    num_id = word_val(num_pr.find("./w:numId", NS_DOCX))
    ilvl = word_val(num_pr.find("./w:ilvl", NS_DOCX))
    if num_id is None:
        return None

    try:
        return num_id, int(ilvl or "0")
    except ValueError:
        return num_id, 0


def read_docx_styles(zip_file):
    styles = {}
    if "word/styles.xml" not in zip_file.namelist():
        return styles

    root = ET.fromstring(zip_file.read("word/styles.xml"))
    for style in root.findall("w:style", NS_DOCX):
        style_id = style.get(f"{W}styleId")
        if not style_id:
            continue

        name = word_val(style.find("w:name", NS_DOCX)) or ""
        outline = style.find(".//w:outlineLvl", NS_DOCX)
        level = None

        if outline is not None and word_val(outline) is not None:
            try:
                level = int(word_val(outline)) + 1
            except ValueError:
                level = None

        lower_name = name.lower().replace(" ", "")
        lower_id = style_id.lower().replace(" ", "")
        for prefix in ("heading", "标题"):
            for number in range(1, 10):
                if lower_name in (f"{prefix}{number}", f"{prefix}{number}.") or lower_id in (
                    f"{prefix}{number}",
                    f"{prefix}{number}.",
                ):
                    level = number

        if level is not None:
            styles[style_id] = min(max(level, 1), 6)

    return styles


def read_docx_numbering(zip_file):
    if "word/numbering.xml" not in zip_file.namelist():
        return {}

    root = ET.fromstring(zip_file.read("word/numbering.xml"))
    abstract_map = {}
    num_map = {}

    for abstract in root.findall("w:abstractNum", NS_DOCX):
        abstract_id = abstract.get(f"{W}abstractNumId")
        if abstract_id is None:
            continue

        levels = {}
        for level in abstract.findall("w:lvl", NS_DOCX):
            ilvl = level.get(f"{W}ilvl")
            fmt = word_val(level.find("w:numFmt", NS_DOCX)) or "bullet"
            start = word_val(level.find("w:start", NS_DOCX)) or "1"
            try:
                start_value = int(start)
            except ValueError:
                start_value = 1

            try:
                levels[int(ilvl or "0")] = {"fmt": fmt, "start": start_value}
            except ValueError:
                levels[0] = {"fmt": fmt, "start": start_value}
        abstract_map[abstract_id] = levels

    for num in root.findall("w:num", NS_DOCX):
        num_id = num.get(f"{W}numId")
        abstract_id = word_val(num.find("w:abstractNumId", NS_DOCX))
        if num_id is not None and abstract_id in abstract_map:
            levels = {
                ilvl: {"fmt": info["fmt"], "start": info["start"]}
                for ilvl, info in abstract_map[abstract_id].items()
            }

            for override in num.findall("w:lvlOverride", NS_DOCX):
                ilvl_raw = override.get(f"{W}ilvl")
                try:
                    ilvl = int(ilvl_raw or "0")
                except ValueError:
                    ilvl = 0

                if ilvl not in levels:
                    levels[ilvl] = {"fmt": "decimal", "start": 1}

                start_override = word_val(override.find("w:startOverride", NS_DOCX))
                if start_override is not None:
                    try:
                        levels[ilvl]["start"] = int(start_override)
                    except ValueError:
                        pass

                level_override = override.find("w:lvl", NS_DOCX)
                if level_override is not None:
                    fmt = word_val(level_override.find("w:numFmt", NS_DOCX))
                    start = word_val(level_override.find("w:start", NS_DOCX))
                    if fmt is not None:
                        levels[ilvl]["fmt"] = fmt
                    if start is not None:
                        try:
                            levels[ilvl]["start"] = int(start)
                        except ValueError:
                            pass

            num_map[num_id] = levels

    return num_map


def paragraph_heading_level(paragraph, styles):
    direct_level = paragraph_outline_level(paragraph)
    if direct_level is not None:
        return min(max(direct_level, 1), 6)

    style_id = paragraph_style_id(paragraph)
    if style_id in styles:
        return styles[style_id]

    return None


def table_cell_text(cell, image_exporter=None):
    chunks = []
    for paragraph in cell.findall(".//w:p", NS_DOCX):
        for item_type, value in paragraph_content_items(paragraph, image_exporter):
            if item_type == "text":
                chunks.append(markdown_escape(value))
            elif item_type == "image":
                chunks.append(value)
    return "<br>".join(chunks)


def table_to_markdown(table, image_exporter=None):
    rows = []
    for row in table.findall("w:tr", NS_DOCX):
        values = [table_cell_text(cell, image_exporter) for cell in row.findall("w:tc", NS_DOCX)]
        if values:
            rows.append(values)

    if not rows:
        return []

    max_cols = max(len(row) for row in rows)
    normalized = [row + [""] * (max_cols - len(row)) for row in rows]
    header = normalized[0]
    body = normalized[1:]

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")

    return lines


def docx_body_to_markdown(zip_file, image_exporter=None):
    styles = read_docx_styles(zip_file)
    numbering = read_docx_numbering(zip_file)
    counters = {}

    root = ET.fromstring(zip_file.read("word/document.xml"))
    body = root.find("w:body", NS_DOCX)
    if body is None:
        return []

    lines = []
    for child in list(body):
        if child.tag == f"{W}p":
            text = paragraph_text(child)
            images = [
                value
                for item_type, value in paragraph_content_items(child, image_exporter)
                if item_type == "image"
            ]
            if not text and not images:
                continue

            heading_level = paragraph_heading_level(child, styles)
            if heading_level is not None and text:
                lines.append(f"{'#' * heading_level} {text}")
                lines.append("")
                for image in images:
                    lines.append(image)
                    lines.append("")
                counters.clear()
                continue

            num_info = paragraph_numbering(child)
            if num_info is not None and text:
                num_id, ilvl = num_info
                level_info = numbering.get(num_id, {}).get(ilvl, {"fmt": "bullet", "start": 1})
                fmt = level_info["fmt"]
                indent = "  " * ilvl

                if fmt == "decimal":
                    key = (num_id, ilvl)
                    if key not in counters:
                        counters[key] = level_info["start"] - 1
                    counters[key] += 1
                    for stale_key in list(counters):
                        if stale_key[0] == num_id and stale_key[1] > ilvl:
                            del counters[stale_key]
                    lines.append(f"{indent}{counters[key]}. {text}")
                else:
                    lines.append(f"{indent}- {text}")
                for image in images:
                    lines.append(f"{indent}  {image}")
                continue

            content_lines = paragraph_to_markdown_lines(child, image_exporter)
            for line in content_lines:
                lines.append(line)
                lines.append("")

        elif child.tag == f"{W}tbl":
            table_lines = table_to_markdown(child, image_exporter)
            if table_lines:
                lines.extend(table_lines)
                lines.append("")
            counters.clear()

    return lines


def split_markdown_lines(lines):
    split_lines = []
    for line in lines:
        if line == "":
            split_lines.append("")
        else:
            split_lines.extend(str(line).splitlines() or [""])
    return split_lines


def is_heading(line):
    return re.match(r"^#{1,6}\s+\S", line) is not None


def is_list_item(line):
    return re.match(r"^\s*(?:[-*+]|\d+[.)])\s+\S", line) is not None


def is_table_row(line):
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|")


def is_image(line):
    return re.match(r"^\s*!\[[^\]]*\]\([^)]+\)\s*$", line) is not None


def needs_blank_between(previous, current):
    if not previous or not current:
        return False

    if is_table_row(previous) and is_table_row(current):
        return False

    if is_list_item(previous) and is_list_item(current):
        return False

    if is_heading(previous) or is_heading(current):
        return True

    if is_image(previous) or is_image(current):
        return True

    if is_table_row(previous) != is_table_row(current):
        return True

    if is_list_item(previous) != is_list_item(current):
        return True

    return True


def clean_markdown_lines(lines):
    cleaned = []
    previous_nonblank = ""

    for raw_line in split_markdown_lines(lines):
        line = raw_line.rstrip()

        if not line.strip():
            if cleaned and cleaned[-1] != "":
                cleaned.append("")
            continue

        if cleaned and cleaned[-1] != "" and needs_blank_between(previous_nonblank, line):
            cleaned.append("")

        cleaned.append(line)
        previous_nonblank = line

    while cleaned and cleaned[0] == "":
        cleaned.pop(0)
    while cleaned and cleaned[-1] == "":
        cleaned.pop()

    return cleaned


def convert_docx(path, output_dir):
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{path.stem}.md"
    media_dir = output_dir / f"{path.stem}_media"
    with ZipFile(path) as zip_file:
        image_exporter = DocxImageExporter(zip_file, read_docx_relationships(zip_file), media_dir)
        body_lines = docx_body_to_markdown(zip_file, image_exporter)

    lines = body_lines
    lines = clean_markdown_lines(lines)

    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return [output_path]


def discover_default_files():
    files = []
    for suffix in ("*.xlsx", "*.docx"):
        files.extend(BASE.glob(suffix))
    return files


def output_root_for_folder(folder):
    return folder.parent / "converted" / safe_filename(folder.name)


def output_dir_for_file(path):
    return path.parent / "converted"


def supported_file_paths(folder):
    files = []
    for path in folder.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES:
            files.append(path)
    return sorted(files)


def discover_tasks(argv):
    if not argv:
        return [ConversionTask(path, DEFAULT_OUT_DIR) for path in discover_default_files()]

    tasks = []
    for arg in argv:
        path = Path(arg).resolve()
        if not path.exists():
            tasks.append(ConversionTask(path, DEFAULT_OUT_DIR))
            continue

        if path.is_dir():
            folder_output_root = output_root_for_folder(path)
            for file_path in supported_file_paths(path):
                relative_parent = file_path.parent.relative_to(path)
                output_dir = folder_output_root / relative_parent
                tasks.append(ConversionTask(file_path, output_dir))
            continue

        tasks.append(ConversionTask(path, output_dir_for_file(path)))

    return tasks


def convert_file(path, output_dir):
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return convert_xlsx(path, output_dir)
    if suffix == ".docx":
        return convert_docx(path, output_dir)
    print(f"Skipped unsupported file: {path}")
    return []


def main(argv):
    DEFAULT_OUT_DIR.mkdir(exist_ok=True)

    tasks = discover_tasks(argv)
    existing_tasks = [task for task in tasks if task.source_path.exists()]
    if not existing_tasks:
        print("No .xlsx or .docx files found.")
        return 1

    all_written = []
    had_error = False

    for task in tasks:
        path = task.source_path
        if not path.exists():
            print(f"Missing file: {path}")
            had_error = True
            continue

        try:
            written = convert_file(path, task.output_dir)
            all_written.extend(written)
            if written:
                print(f"Converted: {path.name}")
                for output in written:
                    print(f"  -> {output}")
        except Exception:
            had_error = True
            print(f"ERROR converting: {path}")
            traceback.print_exc()

    print("")
    print(f"Default output folder: {DEFAULT_OUT_DIR}")
    print(f"Files written: {len(all_written)}")
    return 1 if had_error else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
