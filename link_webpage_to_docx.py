#!/usr/bin/env python3
"""Copy hyperlinks from an English web page into a translated DOCX file."""

from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import shutil
import sys
import tempfile
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix if prefix != "rel" else "", uri)


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


@dataclass
class SourceLink:
    index: int
    text: str
    url: str
    context: str


@dataclass
class ParagraphRef:
    index: int
    text: str
    element: ET.Element


@dataclass
class LinkMatch:
    link_index: int
    paragraph_index: int
    target_text: str
    confidence: float
    reason: str = ""


class AnchorExtractor(HTMLParser):
    def __init__(self, base_url: str):
        super().__init__(convert_charrefs=True)
        self.base_url = base_url
        self.links: List[SourceLink] = []
        self._text_parts: List[str] = []
        self._active_href: Optional[str] = None
        self._active_text: List[str] = []

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        if tag.lower() != "a":
            return
        attrs_dict = {key.lower(): value for key, value in attrs}
        href = attrs_dict.get("href")
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
            return
        self._active_href = urllib.parse.urljoin(self.base_url, href)
        self._active_text = []

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() != "a" or not self._active_href:
            return
        anchor_text = normalize_space("".join(self._active_text))
        if anchor_text:
            context = normalize_space(" ".join(self._text_parts[-80:] + [anchor_text]))
            self.links.append(
                SourceLink(
                    index=len(self.links),
                    text=anchor_text,
                    url=self._active_href,
                    context=context[-900:],
                )
            )
        self._text_parts.append(anchor_text)
        self._active_href = None
        self._active_text = []

    def handle_data(self, data: str) -> None:
        text = normalize_space(data)
        if not text:
            return
        if self._active_href:
            self._active_text.append(text)
        else:
            self._text_parts.append(text)


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(value or "")).strip()


def fetch_html(url: str) -> str:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (compatible; WebDocxLinker/1.0)",
            "Accept": "text/html,application/xhtml+xml",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def extract_links(url: str, max_links: Optional[int] = None) -> List[SourceLink]:
    parser = AnchorExtractor(url)
    parser.feed(fetch_html(url))
    unique: Dict[Tuple[str, str], SourceLink] = {}
    for link in parser.links:
        key = (link.text.casefold(), link.url)
        if key not in unique:
            unique[key] = SourceLink(len(unique), link.text, link.url, link.context)
    links = list(unique.values())
    return links[:max_links] if max_links else links


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t")))


def read_docx_paragraphs(docx_path: Path) -> Tuple[ET.ElementTree, List[ParagraphRef]]:
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    paragraphs: List[ParagraphRef] = []
    for paragraph in tree.getroot().iter(qn("w:p")):
        text = normalize_space(paragraph_text(paragraph))
        if text:
            paragraphs.append(ParagraphRef(len(paragraphs), text, paragraph))
    return tree, paragraphs


def compact_paragraphs(paragraphs: Sequence[ParagraphRef], max_chars: int = 36000) -> List[Dict[str, object]]:
    output: List[Dict[str, object]] = []
    used = 0
    for para in paragraphs:
        text = para.text
        item = {"paragraph_index": para.index, "text": text[:900]}
        size = len(json.dumps(item, ensure_ascii=False))
        if used + size > max_chars:
            break
        output.append(item)
        used += size
    return output


def match_links_with_openai(
    links: Sequence[SourceLink],
    paragraphs: Sequence[ParagraphRef],
    model: str,
    api_key: Optional[str] = None,
) -> List[LinkMatch]:
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise SystemExit(
            "openai 패키지가 필요합니다. 먼저 `pip install -r requirements.txt`를 실행하세요."
        ) from exc

    client = OpenAI(api_key=api_key) if api_key else OpenAI()
    link_payload = [
        {
            "link_index": link.index,
            "source_anchor_text": link.text,
            "url": link.url,
            "source_context": link.context,
        }
        for link in links
    ]
    paragraph_payload = compact_paragraphs(paragraphs)
    prompt = {
        "task": (
            "The links come from an English source webpage. The paragraphs come from a Korean "
            "Word translation of that page. For each source link, find the exact Korean substring "
            "in the translated paragraph that should receive the same hyperlink."
        ),
        "rules": [
            "Return only JSON.",
            "Use the exact substring as it appears in the Korean paragraph.",
            "If there is no good match, use paragraph_index -1, target_text empty, confidence 0.",
            "Prefer linked noun phrases, product names, organization names, article titles, or exact translated terms.",
            "Do not invent target text that is absent from the paragraph.",
        ],
        "links": link_payload,
        "paragraphs": paragraph_payload,
        "output_schema": {
            "matches": [
                {
                    "link_index": "integer",
                    "paragraph_index": "integer",
                    "target_text": "string",
                    "confidence": "number from 0 to 1",
                    "reason": "short Korean explanation",
                }
            ]
        },
    }
    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "system",
                "content": "You align English webpage hyperlinks to a Korean DOCX translation with high precision.",
            },
            {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
        ],
        text={"format": {"type": "json_object"}},
    )
    payload = json.loads(response.output_text)
    matches = []
    for item in payload.get("matches", []):
        try:
            matches.append(
                LinkMatch(
                    link_index=int(item["link_index"]),
                    paragraph_index=int(item["paragraph_index"]),
                    target_text=str(item.get("target_text", "")).strip(),
                    confidence=float(item.get("confidence", 0)),
                    reason=str(item.get("reason", "")),
                )
            )
        except (KeyError, TypeError, ValueError):
            continue
    return matches


def copy_run_with_text(run: ET.Element, text: str, hyperlink_style: bool = False) -> ET.Element:
    new_run = ET.Element(qn("w:r"))
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
        new_run.append(deepcopy_element(rpr))
    if hyperlink_style:
        new_rpr = new_run.find(qn("w:rPr"))
        if new_rpr is None:
            new_rpr = ET.Element(qn("w:rPr"))
            new_run.insert(0, new_rpr)
        for old_style in list(new_rpr.findall(qn("w:rStyle"))):
            new_rpr.remove(old_style)
        style = ET.Element(qn("w:rStyle"), {qn("w:val"): "Hyperlink"})
        new_rpr.insert(0, style)
    text_node = ET.SubElement(new_run, qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return new_run


def deepcopy_element(element: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def relationship_id(rels_root: ET.Element, url: str) -> str:
    existing_ids = []
    for rel in rels_root.findall("rel:Relationship", NS):
        existing_ids.append(rel.attrib.get("Id", ""))
        if (
            rel.attrib.get("Type")
            == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            and rel.attrib.get("Target") == url
        ):
            return rel.attrib["Id"]
    max_id = 0
    for rid in existing_ids:
        match = re.match(r"rId(\d+)$", rid or "")
        if match:
            max_id = max(max_id, int(match.group(1)))
    new_id = f"rId{max_id + 1}"
    ET.SubElement(
        rels_root,
        f"{{{NS['rel']}}}Relationship",
        {
            "Id": new_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            "Target": url,
            "TargetMode": "External",
        },
    )
    return new_id


def text_positions(paragraph: ET.Element) -> Tuple[str, List[Tuple[int, int, ET.Element]]]:
    pieces: List[str] = []
    positions: List[Tuple[int, int, ET.Element]] = []
    cursor = 0
    for run in paragraph.findall(qn("w:r")):
        text = "".join(node.text or "" for node in run.iter(qn("w:t")))
        if not text:
            continue
        start = cursor
        cursor += len(text)
        pieces.append(text)
        positions.append((start, cursor, run))
    return "".join(pieces), positions


def wrap_hyperlink(paragraph: ET.Element, target_text: str, url: str, rels_root: ET.Element) -> bool:
    full_text, positions = text_positions(paragraph)
    start = full_text.find(target_text)
    if start < 0:
        start = normalize_space(full_text).find(normalize_space(target_text))
        if start < 0:
            return False
    end = start + len(target_text)
    rid = relationship_id(rels_root, url)
    new_children: List[ET.Element] = []
    hyperlink = ET.Element(qn("w:hyperlink"), {qn("r:id"): rid})
    inserted_any = False

    for child in list(paragraph):
        if child.tag != qn("w:r"):
            new_children.append(deepcopy_element(child))
            continue
        run_text = "".join(node.text or "" for node in child.iter(qn("w:t")))
        run_pos = next((pos for pos in positions if pos[2] is child), None)
        if not run_text or run_pos is None:
            new_children.append(deepcopy_element(child))
            continue
        run_start, run_end, _ = run_pos
        if run_end <= start or run_start >= end:
            new_children.append(deepcopy_element(child))
            continue
        before = run_text[: max(0, start - run_start)]
        linked = run_text[max(0, start - run_start) : min(run_end, end) - run_start]
        after = run_text[min(run_end, end) - run_start :]
        if before:
            new_children.append(copy_run_with_text(child, before))
        if linked:
            hyperlink.append(copy_run_with_text(child, linked, hyperlink_style=True))
            inserted_any = True
        if after:
            new_children.append(copy_run_with_text(child, after))

    if not inserted_any:
        return False
    rebuilt: List[ET.Element] = []
    hyperlink_added = False
    for child in new_children:
        if not hyperlink_added and child.tag == qn("w:r"):
            child_text = "".join(node.text or "" for node in child.iter(qn("w:t")))
            prefix = full_text[:start]
            current_prefix = "".join(
                "".join(node.text or "" for node in item.iter(qn("w:t"))) for item in rebuilt
            )
            if len(current_prefix) >= len(prefix):
                rebuilt.append(hyperlink)
                hyperlink_added = True
        rebuilt.append(child)
    if not hyperlink_added:
        rebuilt.append(hyperlink)
    paragraph[:] = rebuilt
    return True


def write_docx(
    source_docx: Path,
    output_docx: Path,
    document_tree: ET.ElementTree,
    rels_root: ET.Element,
) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        with zipfile.ZipFile(source_docx) as archive:
            archive.extractall(temp_path)
        document_tree.write(
            temp_path / "word/document.xml",
            encoding="utf-8",
            xml_declaration=True,
        )
        rels_dir = temp_path / "word/_rels"
        rels_dir.mkdir(parents=True, exist_ok=True)
        ET.ElementTree(rels_root).write(
            rels_dir / "document.xml.rels",
            encoding="utf-8",
            xml_declaration=True,
        )
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", temp_path)
        zip_output = output_docx.with_suffix(".zip")
        if output_docx.exists():
            output_docx.unlink()
        zip_output.rename(output_docx)


def read_relationships(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as archive:
        try:
            xml_bytes = archive.read("word/_rels/document.xml.rels")
        except KeyError:
            return ET.Element(f"{{{NS['rel']}}}Relationships")
    return ET.fromstring(xml_bytes)


def save_report(path: Path, rows: Iterable[Dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = list(rows)
    fieldnames = [
        "link_index",
        "source_text",
        "url",
        "paragraph_index",
        "target_text",
        "confidence",
        "inserted",
        "reason",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def process_docx_links(
    url: str,
    docx_path: Path,
    output_docx: Path,
    model: str = "gpt-4.1-mini",
    min_confidence: float = 0.62,
    max_links: Optional[int] = None,
    report_path: Optional[Path] = None,
    dry_run: bool = False,
    api_key: Optional[str] = None,
) -> Dict[str, object]:
    links = extract_links(url, max_links)
    if not links:
        raise RuntimeError("웹페이지에서 처리할 링크를 찾지 못했습니다.")

    document_tree, paragraphs = read_docx_paragraphs(docx_path)
    rels_root = read_relationships(docx_path)
    matches = match_links_with_openai(links, paragraphs, model, api_key=api_key)

    links_by_index = {link.index: link for link in links}
    paragraphs_by_index = {paragraph.index: paragraph for paragraph in paragraphs}
    report_rows: List[Dict[str, object]] = []
    inserted_count = 0

    for match in matches:
        link = links_by_index.get(match.link_index)
        paragraph = paragraphs_by_index.get(match.paragraph_index)
        inserted = False
        if (
            link
            and paragraph
            and match.target_text
            and match.confidence >= min_confidence
            and not dry_run
        ):
            inserted = wrap_hyperlink(paragraph.element, match.target_text, link.url, rels_root)
            inserted_count += int(inserted)
        report_rows.append(
            {
                "link_index": match.link_index,
                "source_text": link.text if link else "",
                "url": link.url if link else "",
                "paragraph_index": match.paragraph_index,
                "target_text": match.target_text,
                "confidence": f"{match.confidence:.2f}",
                "inserted": inserted,
                "reason": match.reason,
            }
        )

    if report_path:
        save_report(report_path, report_rows)
    if not dry_run:
        write_docx(docx_path, output_docx, document_tree, rels_root)

    return {
        "link_count": len(links),
        "match_count": len(matches),
        "inserted_count": inserted_count,
        "report_rows": report_rows,
        "output_docx": output_docx,
        "report_path": report_path,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True, help="원문 영문 웹페이지 URL")
    parser.add_argument("--docx", required=True, type=Path, help="한국어 번역 Word 문서")
    parser.add_argument("--out", required=True, type=Path, help="링크가 삽입된 Word 문서 출력 경로")
    parser.add_argument("--model", default="gpt-4.1-mini", help="OpenAI 모델명")
    parser.add_argument("--min-confidence", type=float, default=0.62, help="삽입할 최소 매칭 신뢰도")
    parser.add_argument("--max-links", type=int, default=None, help="처리할 최대 링크 수")
    parser.add_argument("--dry-run", action="store_true", help="수정 없이 매칭 결과만 출력")
    parser.add_argument("--report", type=Path, default=None, help="CSV 리포트 저장 경로")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not os.getenv("OPENAI_API_KEY"):
        print("OPENAI_API_KEY 환경변수가 필요합니다.", file=sys.stderr)
        return 2
    if not args.docx.exists():
        print(f"DOCX 파일을 찾을 수 없습니다: {args.docx}", file=sys.stderr)
        return 2

    result = process_docx_links(
        url=args.url,
        docx_path=args.docx,
        output_docx=args.out,
        model=args.model,
        min_confidence=args.min_confidence,
        max_links=args.max_links,
        report_path=args.report,
        dry_run=args.dry_run,
    )

    print(f"추출 링크: {result['link_count']}")
    print(f"매칭 결과: {result['match_count']}")
    print(f"삽입 완료: {result['inserted_count']}")
    if args.report:
        print(f"리포트: {args.report}")
    if not args.dry_run:
        print(f"출력 파일: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
