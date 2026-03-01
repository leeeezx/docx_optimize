#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查 Word 文档中是否存在 MathType OLE 对象。

支持格式：
1. .docx/.docm/.dotx/.dotm：通过 OpenXML（zip + xml）静态检查
2. .doc：通过 Word COM 检查（需要 Windows + 本机安装 Word + pywin32）

退出码约定：
0 = 未发现 MathType OLE 对象
1 = 发现 MathType OLE 对象
2 = 执行出错（参数错误、文件打不开、环境不满足等）
"""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable


# 常见 MathType OLE ProgID：Equation.DSMT4 / Equation.DSMT6 等
MATHTYPE_PATTERNS = (
    re.compile(r"equation\.dsmt\d*", re.IGNORECASE),
    re.compile(r"mathtype", re.IGNORECASE),
)


@dataclass
class CheckResult:
    found: bool
    count: int
    mode: str
    details: list[str] = field(default_factory=list)


def _is_mathtype_text(text: str | None) -> bool:
    if not text:
        return False
    return any(p.search(text) for p in MATHTYPE_PATTERNS)


def _scan_xml_for_mathtype_ole_objects(
    xml_bytes: bytes, part_name: str, detail_limit: int = 20
) -> tuple[int, list[str]]:
    hits = 0
    details: list[str] = []

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        # 极少数 XML 可能不规范，降级为保守匹配
        text = xml_bytes.decode("utf-8", errors="ignore")
        if _is_mathtype_text(text) and "oleobject" in text.lower():
            return 1, [f"{part_name}: 文本匹配到疑似 MathType OLEObject"]
        return 0, []

    for elem in root.iter():
        # 只统计 OLEObject 节点，避免将同一对象在不同元数据中重复计数
        tag_name = elem.tag.rsplit("}", 1)[-1].lower()
        if tag_name != "oleobject":
            continue

        prog_id = ""
        for key, value in elem.attrib.items():
            if key.lower().endswith("progid"):
                prog_id = value
                break

        if not prog_id:
            # 某些文档可能不标准，兜底检查当前节点属性文本
            prog_id = " ".join(elem.attrib.values())

        if _is_mathtype_text(prog_id):
            hits += 1
            if len(details) < detail_limit:
                details.append(f"{part_name}: 标签={elem.tag} ProgID={prog_id}")

    return hits, details


def _iter_openxml_parts(names: Iterable[str]) -> Iterable[str]:
    for name in names:
        low = name.lower()
        if low.startswith("word/") and low.endswith(".xml"):
            yield name


def check_openxml_document(path: Path) -> CheckResult:
    count = 0
    details: list[str] = []

    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()

        # 1) 以 XML 中的 OLEObject 为主计数（这是对象级计数）
        for part in _iter_openxml_parts(names):
            xml_bytes = zf.read(part)
            hit_count, hit_details = _scan_xml_for_mathtype_ole_objects(xml_bytes, part)
            count += hit_count
            details.extend(hit_details)

        # 2) 若 XML 未识别到对象，再以 embeddings 做兜底估计，避免重复计数
        if count == 0:
            for name in names:
                low = name.lower()
                if low.startswith("word/embeddings/"):
                    data = zf.read(name)
                    if _is_mathtype_text(data.decode("latin-1", errors="ignore")):
                        count += 1
                        if len(details) < 20:
                            details.append(f"{name}: 二进制内容匹配到 MathType 特征（兜底）")

    return CheckResult(found=count > 0, count=count, mode="openxml", details=details)


def _safe_get_progid(shape_obj) -> str:
    try:
        prog_id = str(shape_obj.OLEFormat.ProgID)
    except Exception:
        return ""
    return prog_id


def check_word_with_com(path: Path) -> CheckResult:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("缺少 pywin32，无法执行 .doc 的 COM 检查") from exc

    count = 0
    details: list[str] = []
    word = None
    doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(
            str(path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )

        for i in range(1, doc.InlineShapes.Count + 1):
            prog_id = _safe_get_progid(doc.InlineShapes.Item(i))
            if _is_mathtype_text(prog_id):
                count += 1
                if len(details) < 20:
                    details.append(f"InlineShape[{i}] ProgID={prog_id}")

        for i in range(1, doc.Shapes.Count + 1):
            prog_id = _safe_get_progid(doc.Shapes.Item(i))
            if _is_mathtype_text(prog_id):
                count += 1
                if len(details) < 20:
                    details.append(f"Shape[{i}] ProgID={prog_id}")

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    return CheckResult(found=count > 0, count=count, mode="com", details=details)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="检查 Word 文档中是否存在 MathType OLE 对象",
    )
    parser.add_argument(
        "file",
        help="Word 文件路径（支持 .docx/.docm/.dotx/.dotm/.doc）",
    )
    parser.add_argument(
        "--show-details",
        action="store_true",
        help="输出匹配详情（最多 20 条）",
    )
    parser.add_argument(
        "--force-com",
        action="store_true",
        help="对 .docx/.docm 也强制使用 Word COM 检查（需 pywin32 + Word）",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists() or not path.is_file():
        print(f"错误：文件不存在 -> {path}", file=sys.stderr)
        return 2

    ext = path.suffix.lower()
    try:
        if args.force_com:
            result = check_word_with_com(path)
        elif ext in {".docx", ".docm", ".dotx", ".dotm"}:
            result = check_openxml_document(path)
        elif ext == ".doc":
            result = check_word_with_com(path)
        else:
            print(f"错误：不支持的文件类型 -> {ext}", file=sys.stderr)
            return 2
    except Exception as exc:
        print(f"错误：检查失败，原因：{exc}", file=sys.stderr)
        return 2

    if result.found:
        print(f"FOUND: 检测到 MathType OLE 对象，数量={result.count}，模式={result.mode}")
        if args.show_details and result.details:
            for item in result.details:
                print(f"- {item}")
        return 1

    print(f"NOT_FOUND: 未检测到 MathType OLE 对象，模式={result.mode}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
