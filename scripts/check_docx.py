#!/usr/bin/env python3
"""云南财经大学本科毕业论文格式自检（离线、只读）。

用法：
    python scripts/check_docx.py <论文.docx> [--profile PROFILE] [--json]
                                 [--only rule.id,rule.id] [--list-rules] [--list-profiles]

数值唯一来源为 references/rules.json，每个 profile 一份口径：

* ``zhonghua-cs-2023``（默认）—— 云南财经大学中华职业学院 计算机科学与技术
  （高职本科）2023 级 / 2027 届专业规范；
* ``school-436`` —— 校教发〔2024〕436 号全校口径。

检查项覆盖页面设置、页码分段、页眉、样式、标题编号、题注、目录、摘要、参考文献、
AI 工具使用声明、装订顺序等可机检内容。退出码：0 表示无 FAIL，1 表示存在 FAIL。

检查读取 OOXML 与样式定义，不做渲染。页眉串节、目录点线、图表跨页断开等版式问题
仍须导出 PDF 转图逐页核对；脚本结论与渲染结果冲突时，以渲染结果为准。

少数规则只作参考、不做机检（如字号依赖样式继承、须肉眼判断的项），--list-rules
仍会列出它们。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
R_ID = R + "id"

CHINESE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
CAPTION = re.compile(r"^(图|表)(\d+)-(\d+)\s*(.*)$")
TRAILING_PUNCT = "。．，；：、,;:.!？?"
TWIPS_PER_CM = 566.929


def w(tag: str) -> str:
    return W + tag


def norm(text: str) -> str:
    """去掉所有空白（含全角空格），用于题头等短文本匹配。"""
    return re.sub(r"[\s\u3000]+", "", text or "")


def attr(el, name, default=None):
    return el.get(w(name), default) if el is not None else default


class Docx:
    """按需读取 .docx 内的结构与样式事实。"""

    def __init__(self, path: Path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())
        self.rels = self._rels("word/_rels/document.xml.rels")
        self.styles = self._parse("word/styles.xml")
        self.numbering = self._parse("word/numbering.xml")
        self.doc = self._parse("word/document.xml")
        self.settings = self._parse("word/settings.xml")
        self.style_index = self._style_index()
        self.body = self.doc.find(w("body"))
        self.blocks = self._blocks()
        self.sections = self._sections()

    # -- low level ----------------------------------------------------------

    def _parse(self, name):
        if name not in self.names:
            return None
        return ET.fromstring(self.zip.read(name))

    def _rels(self, name):
        out = {}
        root = self._parse(name)
        if root is None:
            return out
        for rel in root:
            out[rel.get("Id")] = rel.get("Target")
        return out

    def part(self, rel_id):
        target = self.rels.get(rel_id)
        if not target:
            return None, None
        target = target.replace("../", "")
        name = target if target.startswith("word/") else "word/" + target
        return self._parse(name), name

    # -- styles -------------------------------------------------------------

    def _style_index(self):
        index = {}
        if self.styles is None:
            return index
        for st in self.styles.findall(w("style")):
            index[st.get(w("styleId"))] = st
        return index

    def style_by_name(self, name):
        if self.styles is None:
            return None
        for st in self.styles.findall(w("style")):
            nm = st.find(w("name"))
            if nm is not None and nm.get(w("val")) == name:
                return st
        return None

    def _doc_defaults_rpr(self):
        if self.styles is None:
            return None
        dd = self.styles.find(w("docDefaults"))
        if dd is None:
            return None
        rpr = dd.find(w("rPrDefault"))
        return rpr.find(w("rPr")) if rpr is not None else None

    def _chain(self, style_id):
        chain, seen = [], set()
        while style_id and style_id in self.style_index and style_id not in seen:
            seen.add(style_id)
            st = self.style_index[style_id]
            chain.append(st)
            base = st.find(w("basedOn"))
            style_id = attr(base, "val")
        return list(reversed(chain))

    def resolve_rpr(self, style_id):
        props = {}
        sources = []
        defaults = self._doc_defaults_rpr()
        if defaults is not None:
            sources.append(defaults)
        for st in self._chain(style_id):
            rpr = st.find(w("rPr"))
            if rpr is not None:
                sources.append(rpr)
        for rpr in sources:
            fonts = rpr.find(w("rFonts"))
            if fonts is not None:
                bucket = props.setdefault("fonts", {})
                for key in ("ascii", "eastAsia", "hAnsi"):
                    if attr(fonts, key):
                        bucket[key] = attr(fonts, key)
            sz = rpr.find(w("sz"))
            if sz is not None:
                props["sz"] = int(attr(sz, "val"))
            b = rpr.find(w("b"))
            if b is not None:
                props["bold"] = attr(b, "val") not in ("0", "false")
        return props

    def resolve_ppr(self, style_id):
        props = {}
        for st in self._chain(style_id):
            ppr = st.find(w("pPr"))
            if ppr is None:
                continue
            sp = ppr.find(w("spacing"))
            if sp is not None:
                for key in ("line", "lineRule", "before", "after"):
                    if attr(sp, key) is not None:
                        props[key] = attr(sp, key)
            ind = ppr.find(w("ind"))
            if ind is not None:
                for key in ("firstLine", "firstLineChars", "left", "leftChars"):
                    if attr(ind, key) is not None:
                        props[key] = attr(ind, key)
            ol = ppr.find(w("outlineLvl"))
            if ol is not None:
                props["outlineLvl"] = int(attr(ol, "val"))
            numpr = ppr.find(w("numPr"))
            if numpr is not None:
                nid = numpr.find(w("numId"))
                ilvl = numpr.find(w("ilvl"))
                if nid is not None:
                    props["numId"] = attr(nid, "val")
                props["ilvl"] = attr(ilvl, "val", "0") if ilvl is not None else props.get("ilvl", "0")
        return props

    def abstract_levels(self, num_id):
        if self.numbering is None or num_id is None:
            return {}
        abstract_id = None
        for num in self.numbering.findall(w("num")):
            if num.get(w("numId")) == str(num_id):
                node = num.find(w("abstractNumId"))
                abstract_id = attr(node, "val")
                break
        if abstract_id is None:
            return {}
        for abstract in self.numbering.findall(w("abstractNum")):
            if abstract.get(w("abstractNumId")) != abstract_id:
                continue
            levels = {}
            for lvl in abstract.findall(w("lvl")):
                fmt = lvl.find(w("numFmt"))
                txt = lvl.find(w("lvlText"))
                suff = lvl.find(w("suff"))
                levels[lvl.get(w("ilvl"))] = {
                    "fmt": attr(fmt, "val"),
                    "text": attr(txt, "val"),
                    "suff": attr(suff, "val"),
                }
            return levels
        return {}

    # -- body ---------------------------------------------------------------

    @staticmethod
    def para_text(p):
        return "".join(t.text or "" for t in p.iter(w("t")))

    @staticmethod
    def para_style(p):
        ppr = p.find(w("pPr"))
        if ppr is None:
            return None
        st = ppr.find(w("pStyle"))
        return attr(st, "val")

    def _blocks(self):
        blocks = []
        for child in self.body:
            if child.tag == w("p"):
                ppr = child.find(w("pPr"))
                blocks.append({
                    "kind": "p",
                    "el": child,
                    "text": self.para_text(child),
                    "style": self.para_style(child),
                    "has_image": child.find(".//" + w("drawing")) is not None
                    or child.find(".//" + w("pict")) is not None
                    or child.find(".//" + w("object")) is not None,
                    "sectPr": ppr.find(w("sectPr")) if ppr is not None else None,
                })
            elif child.tag == w("tbl"):
                blocks.append({"kind": "tbl", "el": child, "text": "", "style": None,
                               "has_image": False, "sectPr": None})
        return blocks

    def _sections(self):
        sections, current = [], []
        for i, blk in enumerate(self.blocks):
            current.append(i)
            if blk.get("sectPr") is not None:
                sections.append({"sectPr": blk["sectPr"], "blocks": current})
                current = []
        sections.append({"sectPr": self.body.find(w("sectPr")), "blocks": current})
        out = []
        for n, sec in enumerate(sections, 1):
            if sec["sectPr"] is None:
                continue
            sec["index"] = n
            out.append(sec)
        return out

    def heading_level(self, style_id):
        st = self.style_index.get(style_id)
        if st is None:
            return None
        nm = st.find(w("name"))
        name = (attr(nm, "val") or "").lower()
        m = re.fullmatch(r"heading\s*([1-9])", name)
        if m:
            return int(m.group(1))
        if name == "heading":
            return self.heading_level(attr(st.find(w("basedOn")), "val"))
        return None


class Report:
    def __init__(self, rules):
        self.rules = rules
        self.findings = []

    def add(self, rule_id, status, message, where=""):
        rule = self.rules.get(rule_id, {})
        self.findings.append({
            "id": rule_id,
            "label": rule.get("label", rule_id),
            "status": status,
            "message": message,
            "where": where,
            "clause": rule.get("clause", ""),
        })

    def ok(self, rule_id, message="符合", where=""):
        self.add(rule_id, "PASS", message, where)

    def fail(self, rule_id, message, where=""):
        self.add(rule_id, "FAIL", message, where)

    def warn(self, rule_id, message, where=""):
        self.add(rule_id, "WARN", message, where)

    def info(self, rule_id, message, where=""):
        self.add(rule_id, "INFO", message, where)


# ------------------------------------------------------------------- helpers


def body_section_index(doc: Docx):
    """正文起始节 = 含第一个一级标题的节。"""
    for sec in doc.sections:
        for i in sec["blocks"]:
            blk = doc.blocks[i]
            if blk["kind"] == "p" and doc.heading_level(blk["style"]) == 1:
                return sec["index"]
    return len(doc.sections)


def find_para(doc: Docx, matches):
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] == "p" and matches(blk["text"].strip()):
            return i, blk
    return None, None


def para_jc(blk):
    ppr = blk["el"].find(w("pPr"))
    jc = ppr.find(w("jc")) if ppr is not None else None
    return attr(jc, "val") if jc is not None else None


def effective_run_props(doc: Docx, blk):
    """显式 run 属性优先，其次段落样式解析结果。"""
    props = {}
    style_id = blk["style"]
    if style_id:
        style_props = doc.resolve_rpr(style_id)
        props.update({"sz": style_props.get("sz"), "bold": style_props.get("bold", False)})
        props.update(style_props.get("fonts", {}))
    for item in run_items(blk):
        if not item["text"].strip():
            continue
        for key in ("eastAsia", "ascii", "sz", "bold"):
            if item.get(key) is not None:
                props[key] = item[key]
    return props


def run_items(blk):
    """按出现顺序列出段落内每个 run 的显式格式（未显式设置则缺键）。"""
    items = []
    for r in blk["el"].iter(w("r")):
        info = {"text": "".join(t.text or "" for t in r.iter(w("t")))}
        rpr = r.find(w("rPr"))
        if rpr is not None:
            fonts = rpr.find(w("rFonts"))
            if fonts is not None:
                for key in ("eastAsia", "ascii", "hAnsi"):
                    if attr(fonts, key):
                        info[key] = attr(fonts, key)
            sz = rpr.find(w("sz"))
            if sz is not None:
                info["sz"] = int(attr(sz, "val"))
            b = rpr.find(w("b"))
            if b is not None:
                info["bold"] = attr(b, "val") not in ("0", "false")
        items.append(info)
    return items


# ------------------------------------------------------------------- checks


def check_sections(doc: Docx, rep: Report, rules):
    size = rules["page.size"]["value"]
    tol = rules["page.size"].get("tolerance", 20)
    margin = rules["page.margin"]["value"]
    header_cm = rules["page.headerDistance"]["value"]["header"]
    footer_cm = rules["page.footerDistance"]["value"]["footer"]
    for sec in doc.sections:
        sp = sec["sectPr"]
        where = f"第 {sec['index']} 节"
        pgsz = sp.find(w("pgSz"))
        if pgsz is None:
            rep.fail("page.size", "缺少 pgSz", where)
        else:
            gw, gh = int(attr(pgsz, "w")), int(attr(pgsz, "h"))
            if abs(gw - size["w"]) > tol or abs(gh - size["h"]) > tol:
                rep.fail("page.size", f"{gw}×{gh} twips = {gw / TWIPS_PER_CM:.2f}×{gh / TWIPS_PER_CM:.2f} cm", where)
            else:
                rep.ok("page.size", f"A4 {gw}×{gh} twips", where)
        mar = sp.find(w("pgMar"))
        if mar is None:
            rep.fail("page.margin", "缺少 pgMar", where)
            continue
        bad = [f"{key}={attr(mar, key)}" for key in ("top", "bottom", "left", "right")
               if attr(mar, key) is None or abs(int(attr(mar, key)) - margin[key]) > tol]
        if bad:
            rep.fail("page.margin",
                     "偏差 " + ", ".join(bad)
                     + f"（应 top/bottom={margin['top']}、left/right={margin['left']} twips）", where)
        else:
            rep.ok("page.margin", f"上下 {margin['top']}、左右 {margin['left']} twips", where)
        hv, fv = attr(mar, "header"), attr(mar, "footer")
        if hv is None or abs(int(hv) - header_cm) > tol:
            rep.fail("page.headerDistance", f"header={hv} twips（应 851 ≈1.5cm）", where)
        if fv is None or abs(int(fv) - footer_cm) > tol:
            rep.fail("page.footerDistance", f"footer={fv} twips（应 992 ≈1.75cm）", where)


def check_page_numbers(doc: Docx, rep: Report, rules):
    body_at = body_section_index(doc)
    for sec in doc.sections:
        sp = sec["sectPr"]
        where = f"第 {sec['index']} 节"
        pgnum = sp.find(w("pgNumType"))
        fmt = attr(pgnum, "fmt") if pgnum is not None else None
        start = attr(pgnum, "start") if pgnum is not None else None
        has_field, centered = False, None
        for ref in sp.findall(w("footerReference")):
            part, _ = doc.part(ref.get(R_ID))
            if part is None:
                continue
            if any("PAGE" in (i.text or "") for i in part.iter(w("instrText"))):
                has_field = True
                for p in part.iter(w("p")):
                    ppr = p.find(w("pPr"))
                    jc = ppr.find(w("jc")) if ppr is not None else None
                    if jc is not None:
                        centered = attr(jc, "val")
                    break
        if not has_field:
            rep.info("page.numbering", "该节无页码域（封面/扉页允许）", where)
            continue
        if sec["index"] < body_at:
            if fmt != "upperRoman":
                rep.fail("page.numbering", f"前置部分页码格式 fmt={fmt}（应 upperRoman）", where)
            else:
                rep.ok("page.numbering", "前置部分用大写罗马数字", where)
        else:
            if fmt == "upperRoman":
                rep.fail("page.numbering", "正文及以后仍为罗马数字", where)
            elif sec["index"] == body_at and str(start) != "1":
                rep.fail("page.numbering", f"正文起始页码 start={start}（应 1，从 1 重新开始）", where)
            else:
                rep.ok("page.numbering", f"阿拉伯数字 fmt={fmt or 'decimal'} start={start}", where)
        if centered is None:
            rep.warn("page.numberingCentered", "页脚未显式设置居中对齐，需渲染确认", where)
        elif centered != "center":
            rep.warn("page.numberingCentered", f"页脚对齐 jc={centered}（应 center）", where)
        else:
            rep.ok("page.numberingCentered", "页码居中", where)


def check_headers(doc: Docx, rep: Report, rules):
    body_at = body_section_index(doc)
    spec = rules["header.content"]["value"]
    mode = spec.get("mode", "styleRef")
    want_jc = spec.get("jc", "center")
    for sec in doc.sections:
        sp = sec["sectPr"]
        where = f"第 {sec['index']} 节"
        refs = sp.findall(w("headerReference"))
        if not refs:
            if sec["index"] < body_at:
                rep.info("header.everyPage", "该节无页眉（封面/扉页允许）", where)
            else:
                rep.warn("header.everyPage", "正文节没有页眉引用，会继承上一节页眉", where)
            continue
        texts, has_field, deviations = [], False, set()
        for ref in refs:
            part, _ = doc.part(ref.get(R_ID))
            if part is None:
                continue
            texts.append("".join(t.text or "" for t in part.iter(w("t"))).strip())
            field = spec.get("field")
            if field and any(field in (i.text or "") for i in part.iter(w("instrText"))):
                has_field = True
            for p in part.iter(w("p")):
                ppr = p.find(w("pPr"))
                jc = ppr.find(w("jc")) if ppr is not None else None
                if jc is not None and attr(jc, "val") != want_jc:
                    deviations.add(f"对齐={attr(jc, 'val')}")
                for r in p.iter(w("r")):
                    rpr = r.find(w("rPr"))
                    sz = rpr.find(w("sz")) if rpr is not None else None
                    if sz is not None and spec.get("sz") and attr(sz, "val") != str(spec["sz"]):
                        deviations.add(f"sz={attr(sz, 'val')}")
        if sec["index"] < body_at:
            continue
        if mode == "fixedText":
            want_text = spec.get("text", "")
            hit = [t for t in texts if want_text and want_text in t]
            if hit:
                rep.ok("header.content", f"页眉为固定文字「{want_text}」", where)
            else:
                rep.fail("header.content",
                         f"正文节页眉应统一标注「{want_text}」，实际内容={texts}", where)
        else:
            if has_field:
                rep.ok("header.content", "页眉用 STYLEREF 引用本页一级标题", where)
            else:
                rep.fail("header.content", f'正文节页眉没有 STYLEREF "标题 1" 域，内容={texts}', where)
        if deviations:
            rep.warn("header.content", "页眉格式偏差：" + "、".join(sorted(deviations)), where)


def check_styles(doc: Docx, rep: Report, rules):
    normal = doc.style_by_name("Normal")
    if normal is None:
        rep.fail("style.normal", "找不到 Normal 样式", "")
    else:
        props = doc.resolve_rpr(normal.get(w("styleId")))
        fonts = props.get("fonts", {})
        want = rules["style.normal"]["value"]
        if props.get("sz") == want["sz"] and fonts.get("eastAsia") == want["eastAsia"]:
            rep.ok("style.normal", f"{want['eastAsia']} sz={props.get('sz')}（小 4）", "Normal")
        else:
            rep.fail("style.normal",
                     f"中文字体={fonts.get('eastAsia')} sz={props.get('sz')}（应 宋体 sz=24）", "Normal")
        if fonts.get("ascii") != want["ascii"]:
            rep.warn("style.normal", f"西文字体 ascii={fonts.get('ascii')}（应 Times New Roman）", "Normal")
        pprops = doc.resolve_ppr(normal.get(w("styleId")))
        line, rule = pprops.get("line"), pprops.get("lineRule")
        want_line = rules["page.lineSpacing"]["value"]
        if line == str(want_line["line"]) and rule in (None, want_line["lineRule"]):
            rep.ok("page.lineSpacing", f"line={line} lineRule={rule}（1.5 倍）", "Normal")
        else:
            rep.fail("page.lineSpacing",
                     f"line={line} lineRule={rule}（应 {want_line['line']}/{want_line['lineRule']}）", "Normal")
        if want_line.get("beforeAfterTwips") is not None:
            gaps = {k: pprops.get(k) for k in ("before", "after")}
            nonzero = {k: v for k, v in gaps.items() if v not in (None, "0")}
            if nonzero:
                rep.warn("page.lineSpacing",
                         f"正文段前段后={nonzero} twips（本专业规范要求段前段后 0 行）", "Normal")
        if pprops.get("firstLineChars") == "200":
            rep.ok("style.normalIndent", "首行缩进 2 字符", "Normal")
        else:
            rep.warn("style.normalIndent", f"firstLineChars={pprops.get('firstLineChars')}（中文正文通常 200）", "Normal")
    schemes = rules["heading.font"]["value"]
    active = schemes.get("active") or schemes["default"]
    for level in (1, 2, 3, 4):
        st = doc.style_by_name(f"heading {level}")
        if st is None:
            rep.fail("heading.font", f"缺少 heading {level} 样式", "")
            continue
        props = doc.resolve_rpr(st.get(w("styleId")))
        font = props.get("fonts", {}).get("eastAsia")
        size = props.get("sz")
        bold = bool(props.get("bold"))
        matched = [
            name for name, scheme in schemes["schemes"].items()
            if font == scheme["font"] and size == scheme["sz"][str(level)]
            and bold == bool(scheme.get("bold"))
        ]
        detail = f"{font} sz={size}" + ("加粗" if bold else "未加粗")
        if active in matched:
            rep.ok("heading.font", f"{detail}（{schemes['schemes'][active]['label']}）", f"heading {level}")
        elif matched:
            rep.warn(
                "heading.font",
                f"{detail}，属于「{schemes['schemes'][matched[0]]['label']}」，与默认口径"
                f"「{schemes['schemes'][active]['label']}」不同：请与导师/学院模板核对，"
                f"并用 --heading-scheme {matched[0]} 复检",
                f"heading {level}",
            )
        else:
            expect = "；".join(
                f"{name}={scheme['font']} sz={scheme['sz'][str(level)]}"
                for name, scheme in schemes["schemes"].items()
            )
            rep.fail("heading.font", f"{detail}（与口径不符：{expect}）", f"heading {level}")
        pprops = doc.resolve_ppr(st.get(w("styleId")))
        if pprops.get("outlineLvl") == level - 1:
            rep.ok("heading.outline", f"outlineLvl={level - 1}", f"heading {level}")
        else:
            rep.fail("heading.outline", f"outlineLvl={pprops.get('outlineLvl')}（应 {level - 1}）", f"heading {level}")


def check_heading_numbering(doc: Docx, rep: Report, rules):
    want = rules["heading.numbering"]["value"]
    enforced = want.get("enforce", True)
    if not enforced:
        rep.info("heading.numbering", want.get("note", "本 profile 未强制标题编号形式，跳过"), "")
    for level in (1, 2, 3, 4) if enforced else ():
        st = doc.style_by_name(f"heading {level}")
        if st is None:
            continue
        props = doc.resolve_ppr(st.get(w("styleId")))
        num_id = props.get("numId")
        if num_id is None:
            rep.fail("heading.numbering", "未绑定编号（numId 缺失）", f"heading {level}")
            continue
        levels = doc.abstract_levels(num_id)
        lvl = levels.get(props.get("ilvl", "0"))
        if lvl is None:
            rep.fail("heading.numbering", f"numId={num_id} 无编号定义", f"heading {level}")
            continue
        expect = want[f"lvl{level}"]
        got = (lvl.get("text") or "").replace("%1", "N")
        expected = expect["lvlText"].replace("%1", "N")
        problems = []
        if got != expected:
            problems.append(f"编号形式={lvl.get('text')}（应 {expect['lvlText']}）")
        if lvl.get("suff") != want["suff"]:
            problems.append(f"suff={lvl.get('suff')}（应 nothing，否则目录条目出现空档）")
        if problems:
            rep.fail("heading.numbering", "；".join(problems), f"heading {level}")
        else:
            rep.ok("heading.numbering", f"编号={lvl.get('text')}、suff=nothing", f"heading {level}")
    need = rules["heading.blankLine"]["value"]["minBeforeAfterTwips"]
    for level in (1, 2):
        st = doc.style_by_name(f"heading {level}")
        if st is None:
            continue
        style_id = st.get(w("styleId"))
        props = doc.resolve_ppr(style_id)
        # beforeLines/afterLines 以 1/100 行为单位：100 行即"空一行"
        lines = []
        for node in doc._chain(style_id):
            ppr = node.find(w("pPr"))
            sp = ppr.find(w("spacing")) if ppr is not None else None
            if sp is not None:
                for key in ("beforeLines", "afterLines"):
                    if attr(sp, key) is not None:
                        lines.append(int(attr(sp, key)))
        if lines and min(lines) >= 100:
            rep.ok("heading.blankLine", f"段前/段后各 {min(lines) / 100:.0f} 行", f"heading {level}")
            continue
        before, after = props.get("before"), props.get("after")
        gaps = [int(x) for x in (before, after) if x is not None]
        blank_ok = False
        for i, blk in enumerate(doc.blocks):
            if blk["kind"] != "p" or doc.heading_level(blk["style"]) != level:
                continue
            prev = doc.blocks[i - 1] if i > 0 else None
            nxt = doc.blocks[i + 1] if i + 1 < len(doc.blocks) else None
            if (prev is not None and prev["kind"] == "p" and not prev["text"].strip()
                    and nxt is not None and nxt["kind"] == "p" and not nxt["text"].strip()):
                blank_ok = True
                break
        if blank_ok:
            rep.ok("heading.blankLine", "标题上下各有一个空段落", f"heading {level}")
            continue
        if not gaps:
            rep.warn("heading.blankLine", "未设置段前段后，需确认是否插入了空行", f"heading {level}")
        elif min(gaps) < need:
            rep.warn("heading.blankLine",
                     f"段前/段后={before}/{after} twips，不足一行（应 ≥{need} twips 或改为插入空段落）",
                     f"heading {level}")


def check_front_matter(doc: Docx, rep: Report, rules):
    spec = rules["abstract.heading"]["value"]
    index, blk = find_para(doc, lambda t: norm(t) == "摘要")
    chinese_size = None
    if blk is None:
        rep.fail("abstract.heading", "未找到「摘　要」题头", "")
    else:
        props = effective_run_props(doc, blk)
        chinese_size = props.get("sz")
        problems = []
        if blk["text"].strip() != spec["text"]:
            problems.append(f"题头={blk['text'].strip()!r}（应为 '摘　要'，两字间一个全角空格）")
        if para_jc(blk) != spec["jc"]:
            problems.append(f"对齐={para_jc(blk)}（应 center）")
        if props.get("sz") != spec["sz"]:
            problems.append(f"sz={props.get('sz')}（应 {spec['sz']}）")
        if spec.get("font") and props.get("eastAsia") and props.get("eastAsia") != spec["font"]:
            problems.append(f"字体={props.get('eastAsia')}（应 {spec['font']}）")
        if not props.get("bold"):
            problems.append("未加粗")
        if problems:
            rep.fail("abstract.heading", "；".join(problems), "摘要题头")
        else:
            rep.ok("abstract.heading",
                   f"{spec.get('font', '')} sz={spec['sz']}、加粗、居中、两字间全角空格", "摘要题头")
    if blk is not None:
        chunks = []
        for j in range(index + 1, len(doc.blocks)):
            b = doc.blocks[j]
            if b["kind"] != "p":
                continue
            text = b["text"].strip()
            if text.startswith("关键词") or text.startswith("Keywords"):
                break
            if text:
                chunks.append(text)
        count = len(CHINESE.findall("".join(chunks)))
        low, high = rules["abstract.length"]["value"]["min"], rules["abstract.length"]["value"]["max"]
        if low <= count <= high:
            rep.ok("abstract.length", f"中文 {count} 字", "摘要正文")
        else:
            rep.fail("abstract.length", f"中文 {count} 字（应 {low}—{high}）", "摘要正文")
    _, kw = find_para(doc, lambda t: t.startswith("关键词"))
    if kw is None:
        rep.fail("keywords.count", "未找到关键词段落", "")
    else:
        sep = rules["keywords.count"]["value"].get("separator", "；")
        text = re.sub(r"^关键词[：:　\s]*", "", kw["text"].strip())
        items = [x.strip() for x in re.split(re.escape(sep), text) if x.strip()]
        if len(items) < 2:
            loose = [x for x in re.split(r"[；;、,，\s　]+", text) if x]
            if len(loose) > len(items):
                rep.warn("keywords.count", f"关键词之间应使用「{sep}」分隔（当前用的是其它符号）", "关键词")
                items = loose
        low, high = rules["keywords.count"]["value"]["min"], rules["keywords.count"]["value"]["max"]
        if low <= len(items) <= high:
            rep.ok("keywords.count", f"{len(items)} 个：{sep.join(items)}", "关键词")
        else:
            rep.fail("keywords.count", f"{len(items)} 个（应 {low}—{high}）：{sep.join(items)}", "关键词")
        want = rules["keywords.label"]["value"]
        if not want.get("specified", True):
            rep.info("keywords.label", want.get("note", "本 profile 未规定关键词标签字体，跳过"), "关键词")
        else:
            for item in run_items(kw):
                if "关键词" in item["text"]:
                    if item.get("sz") != want["labelSz"] or item.get("eastAsia") != want["labelFont"]:
                        rep.fail("keywords.label",
                                 f"「关键词」标签 eastAsia={item.get('eastAsia')} sz={item.get('sz')}"
                                 f"（应 {want['labelFont']} sz={want['labelSz']} 即 4 号黑体）", "关键词")
                    else:
                        rep.ok("keywords.label", "「关键词」4 号黑体", "关键词")
                elif item["text"].strip() and item.get("sz") not in (None, want["contentSz"]):
                    rep.warn("keywords.label",
                             f"关键词内容 sz={item.get('sz')}（应 {want['contentSz']} 小 4 宋体，或直接继承正文）", "关键词")
    _, abstract_en = find_para(doc, lambda t: t.strip() == "Abstract")
    if abstract_en is None:
        rep.warn("abstract.englishHeading", "未找到 Abstract 题头", "")
    else:
        props = effective_run_props(doc, abstract_en)
        problems = []
        if para_jc(abstract_en) != rules["abstract.englishHeading"]["value"]["jc"]:
            problems.append(f"对齐={para_jc(abstract_en)}（应 center）")
        if not props.get("bold"):
            problems.append("未加粗")
        if chinese_size is not None and props.get("sz") != chinese_size:
            problems.append(f"sz={props.get('sz')}，与「摘　要」题头 sz={chinese_size} 不一致")
        if problems:
            rep.warn("abstract.englishHeading",
                     "（436 未规定外文题头字号，按一致性要求提示）" + "；".join(problems), "Abstract 题头")
        else:
            rep.ok("abstract.englishHeading", f"居中加粗，字号与摘要题头一致（sz={props.get('sz')}）", "Abstract 题头")
    _, keywords_en = find_para(doc, lambda t: t.strip().startswith("Keywords"))
    if keywords_en is None:
        rep.warn("abstract.english", "未找到 Keywords 行", "")
    else:
        want_en = rules["abstract.english"]["value"]
        for item in run_items(keywords_en):
            if not item["text"].strip():
                continue
            font = item.get("eastAsia") or item.get("ascii")
            if font and font != want_en["font"]:
                rep.fail("abstract.english",
                         f"Keywords 行字体={font}（应 {want_en['font']}）", "Keywords")
            if item.get("sz") not in (None, want_en["sz"]):
                rep.fail("abstract.english",
                         f"Keywords 行 sz={item.get('sz')}（应 {want_en['sz']} 小 4，或继承正文）", "Keywords")
            if "Keywords" in item["text"] and not item.get("bold"):
                rep.fail("abstract.english", "「Keywords」未加粗", "Keywords")
    if abstract_en is not None:
        want_en = rules["abstract.english"]["value"]
        start = doc.blocks.index(abstract_en) + 1
        unusual, checked = set(), 0
        for j in range(start, len(doc.blocks)):
            b = doc.blocks[j]
            if b["kind"] != "p":
                continue
            text = b["text"].strip()
            if not text:
                continue
            if text.startswith("Keywords"):
                break
            if CHINESE.search(text):
                continue
            checked += 1
            for item in run_items(b):
                if not item["text"].strip():
                    continue
                font = item.get("eastAsia") or item.get("ascii")
                if font and font != want_en["font"]:
                    unusual.add(font)
        if unusual and checked:
            rep.fail("abstract.english",
                     f"外文摘要正文出现字体 {'、'.join(sorted(unusual))}（应 {want_en['font']}）", "Abstract 正文")
        elif checked:
            rep.ok("abstract.english", f"外文摘要正文 {checked} 段为 Times New Roman", "Abstract 正文")


def check_toc(doc: Docx, rep: Report, rules):
    spec = rules["toc.heading"]["value"]
    _, blk = find_para(doc, lambda t: norm(t) == "目录")
    if blk is None:
        rep.fail("toc.heading", "未找到「目　录」题头", "")
    else:
        props = effective_run_props(doc, blk)
        problems = []
        if blk["text"].strip() != spec["text"]:
            problems.append(f"题头={blk['text'].strip()!r}（应为 '目　录'）")
        if para_jc(blk) != spec["jc"]:
            problems.append(f"对齐={para_jc(blk)}（应 center）")
        if props.get("sz") != spec["sz"]:
            problems.append(f"sz={props.get('sz')}（应 {spec['sz']} 即 3 号）")
        if spec.get("font") and props.get("eastAsia") and props.get("eastAsia") != spec["font"]:
            problems.append(f"字体={props.get('eastAsia')}（应 {spec['font']}）")
        if not props.get("bold"):
            problems.append("未加粗")
        if problems:
            rep.fail("toc.heading", "；".join(problems), "目录题头")
        else:
            rep.ok("toc.heading",
                   f"{spec.get('font', '')} {spec['sz']}、加粗、居中、两字间全角空格", "目录题头")
    fields = [i.text or "" for i in doc.doc.iter(w("instrText")) if "TOC" in (i.text or "")]
    if not fields:
        rep.warn("toc.levels", "未发现 TOC 域（可能是静态目录）", "")
    else:
        code = fields[0].strip()
        m = re.search(r'\\o\s+"([^"]+)"', code)
        expect = rules["toc.levels"]["value"]["outlineRange"]
        if m and m.group(1) == expect:
            rep.ok("toc.levels", f'TOC \\o "{m.group(1)}"（两级）', "目录域")
        else:
            rep.fail("toc.levels", f'域代码={code}（应包含 \\o "{expect}"）', "目录域")
    entries = [b["text"].strip() for b in doc.blocks
               if b["kind"] == "p" and (b["style"] or "").startswith("TOC") and b["text"].strip()]
    if entries:
        joined = "\n".join(entries)
        missing = [need for need in rules["toc.entries"]["value"]["required"] if need not in joined]
        if missing:
            rep.warn("toc.entries", "目录缺少条目：" + "、".join(missing), "")
        else:
            rep.ok("toc.entries", f"{len(entries)} 条，覆盖前置部分与正文章节", "目录")
    toc1, toc2 = doc.style_by_name("toc 1"), doc.style_by_name("toc 2")
    want = rules["toc.fonts"]["value"]
    if not want.get("specified", True):
        rep.info("toc.fonts", want.get("note", "本 profile 未规定目录条目字体，跳过"), "目录")
    elif toc1 is not None:
        props = doc.resolve_rpr(toc1.get(w("styleId")))
        got = props.get("fonts", {}).get("eastAsia")
        if got == want["toc1Font"]:
            rep.ok("toc.fonts", f"一级条目 {got} sz={props.get('sz')}", "toc 1")
        else:
            rep.fail("toc.fonts", f"一级条目字体={got}（应 {want['toc1Font']}）", "toc 1")
    if toc2 is not None and want.get("specified", True):
        props = doc.resolve_rpr(toc2.get(w("styleId")))
        got = props.get("fonts", {}).get("eastAsia") or "宋体"
        if got == want["toc2Font"]:
            rep.ok("toc.fonts", f"二级条目 {got} sz={props.get('sz')}", "toc 2")


def check_captions(doc: Docx, rep: Report, rules):
    cap_font = rules["caption.font"]["value"]
    figures, tables, captions = {}, {}, []
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] != "p":
            continue
        text = blk["text"].strip()
        m = CAPTION.match(text)
        if not m:
            continue
        kind, chapter, seq = m.group(1), int(m.group(2)), int(m.group(3))
        captions.append((i, kind, chapter, seq, text))
        (figures if kind == "图" else tables).setdefault(chapter, []).append(seq)
        if not m.group(4).strip():
            rep.fail("caption.format", f"题注缺题名：{text}", f"块 {i}")
        elif not re.match(r"^(图|表)\d+-\d+\s", text):
            rep.fail("caption.format", f"编号与题名之间应空一格：{text}", f"块 {i}")
        if text[-1] in TRAILING_PUNCT:
            rep.fail("caption.format", f"题注末尾不应有标点：{text}", f"块 {i}")
        spec = cap_font.get("figure" if kind == "图" else "table", {})
        props = effective_run_props(doc, blk)
        if props.get("sz") != spec.get("sz"):
            rep.fail("caption.font",
                     f"{kind}题 sz={props.get('sz')}（应 {spec.get('sz')}）：{text}", f"块 {i}")
        caption_font = props.get("eastAsia")
        if spec.get("font") and caption_font and caption_font != spec["font"]:
            rep.fail("caption.font",
                     f"{kind}题字体={caption_font}（应 {spec['font']}）：{text}", f"块 {i}")
        if spec.get("bold") and not props.get("bold"):
            rep.fail("caption.font", f"{kind}题未加粗（应 {spec['font']} 加粗）：{text}", f"块 {i}")
        if kind == "图":
            prev = doc.blocks[i - 1] if i > 0 else None
            if prev is None or not (prev["kind"] == "p" and prev["has_image"]):
                rep.fail("caption.position", f"图题上方不是图片：{text}", f"块 {i}")
        else:
            nxt = doc.blocks[i + 1] if i + 1 < len(doc.blocks) else None
            if nxt is None or nxt["kind"] != "tbl":
                rep.fail("caption.position", f"表题下方不是表格：{text}", f"块 {i}")
    for label, store in (("图", figures), ("表", tables)):
        for chapter, seqs in sorted(store.items()):
            expect = list(range(1, len(seqs) + 1))
            if sorted(seqs) != expect:
                rep.fail("caption.sequence", f"{label}{chapter}- 序号 {sorted(seqs)}（应 {expect}）", f"第 {chapter} 章")
    if captions:
        rep.ok("caption.format",
               f"共 {len(captions)} 个题注：图 {sum(len(v) for v in figures.values())}、表 {sum(len(v) for v in tables.values())}", "")
    existing = {f"{k}{c}-{s}" for _, k, c, s, _ in captions}
    first_ref, dangling = {}, set()
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] != "p" or not blk["text"]:
            continue
        for m in re.finditer(r"(图|表)(\d+)-(\d+)", blk["text"]):
            token = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
            if m.group(0).startswith(("图", "表")) and blk["text"].strip().startswith(token):
                continue
            first_ref.setdefault(token, i)
            if token not in existing:
                dangling.add(token)
    for i, kind, chapter, seq, text in captions:
        token = f"{kind}{chapter}-{seq}"
        if token not in first_ref or first_ref[token] > i:
            rep.fail("caption.reference", f"{text} 缺少前置正文引用（应写「如{token}所示」）", f"块 {i}")
    for token in sorted(dangling):
        rep.warn("caption.reference", f"正文引用了不存在的 {token}", "")
    if captions and not any(f["id"] == "caption.reference" and f["status"] == "FAIL" for f in rep.findings):
        rep.ok("caption.reference", "所有题注都有前置正文引用", "")


def check_tables(doc: Docx, rep: Report, rules):
    want = rules["table.borders"]["value"]
    mode = want.get("mode", "threeLine")
    total, clean = 0, 0
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] != "tbl":
            continue
        total += 1
        tblpr = blk["el"].find(w("tblPr"))
        borders = tblpr.find(w("tblBorders")) if tblpr is not None else None
        if borders is None:
            rep.warn("table.borders", "表格未显式声明边框，依赖样式或行内边框，需渲染确认", f"表块 {i}")
            continue
        issues = []
        if mode == "fullGrid":
            for side in want.get("requiredSides", []):
                el = borders.find(w(side))
                if el is None or attr(el, "val") in (None, "none", "nil"):
                    issues.append(f"缺少 {side} 边线（本专业规范要求完整边框）")
        else:
            for side in ("left", "right"):
                el = borders.find(w(side))
                if el is not None and attr(el, "val") not in (None, "none", "nil"):
                    issues.append(f"存在 {side} 边线")
            for side in ("top", "bottom"):
                el = borders.find(w(side))
                sz = attr(el, "sz") if el is not None else None
                if sz is None or int(sz) != want["outerHalfPoints"]:
                    issues.append(f"{side} 边线 sz={sz}（应 {want['outerHalfPoints']} 即 1.5 磅）")
            inside = borders.find(w("insideH"))
            if inside is not None and attr(inside, "sz") and int(attr(inside, "sz")) != want["innerHalfPoints"]:
                issues.append(f"内横线 sz={attr(inside, 'sz')}（应 {want['innerHalfPoints']} 即 0.5 磅）")
        if issues:
            rep.warn("table.borders", "；".join(issues), f"表块 {i}")
        else:
            clean += 1
    if total and clean == total:
        label = "完整边框" if mode == "fullGrid" else "三线表"
        rep.ok("table.borders", f"{total} 个表格符合{label}设置", "")


def check_footnotes(doc: Docx, rep: Report, rules):
    want = rules["note.footnotePerPage"]["value"]
    footnotes = doc._parse("word/footnotes.xml")
    if footnotes is None:
        rep.info("note.footnotePerPage", "未发现 footnotes.xml（尚未使用脚注）", "")
        return
    real = [f for f in footnotes.findall(w("footnote"))
            if f.get(w("type")) is None and int(f.get(w("id"), "0")) > 0]
    if not real:
        rep.info("note.footnotePerPage", "尚未插入任何脚注（题注/注释待补）", "")
        return
    fmt = restart = None
    fpr = doc.settings.find(w("footnotePr")) if doc.settings is not None else None
    if fpr is not None:
        node = fpr.find(w("numFmt"))
        fmt = attr(node, "val") if node is not None else None
        node = fpr.find(w("numRestart"))
        restart = attr(node, "val") if node is not None else None
    problems = []
    if fmt != want["numFmt"]:
        problems.append(f"编号格式={fmt}（应 {want['numFmt']}，即 ①②）")
    if restart != "eachPage":
        problems.append(f"编号重启={restart}（应 eachPage，按页分别编号）")
    if problems:
        rep.warn("note.footnotePerPage", "；".join(problems), "settings.xml")
    else:
        rep.ok("note.footnotePerPage", "脚注 ①②，按页重新编号", "settings.xml")


def check_references(doc: Docx, rep: Report, rules):
    start = end = None
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] != "p":
            continue
        text = blk["text"].strip()
        if start is None and text.rstrip("：:") == "参考文献":
            start = i
        elif start is not None and end is None and re.match(r"^(致谢|附录\s*\d*)", text):
            end = i
    if start is None:
        rep.info("reference.count", "未找到参考文献部分", "")
        return
    end = end or len(doc.blocks)
    entries = []
    for i in range(start + 1, end):
        blk = doc.blocks[i]
        if blk["kind"] == "p" and re.match(r"^\[\d+\]", blk["text"].strip()):
            entries.append(blk["text"].strip())
    network = [e for e in entries if "[EB/OL]" in e or "http" in e]
    solid = [e for e in entries if e not in network]
    low = rules["reference.count"]["value"]["min"]
    if len(solid) >= low:
        rep.ok("reference.count", f"非网络文献 {len(solid)} 条（网络 {len(network)} 条不计入）", "参考文献")
    else:
        rep.fail("reference.count", f"非网络文献 {len(solid)} 条（应 ≥{low}）", "参考文献")
    nums = [int(re.match(r"^\[(\d+)\]", e).group(1)) for e in entries]
    if nums and nums != list(range(1, len(nums) + 1)):
        rep.fail("reference.numbers", f"序号不连续：{nums[:12]}{'…' if len(nums) > 12 else ''}", "参考文献")
    elif nums:
        rep.ok("reference.numbers", f"[1]—[{len(nums)}] 连续", "参考文献")
    years = [int(m.group(0)) for e in entries for m in re.finditer(r"(?:19|20)\d{2}", e)]
    if years:
        rec = rules["reference.recent"]["value"]
        span = rec.get("recentWithinYears", 2)
        rep.info("reference.recent",
                 f"最晚出版年 {max(years)}，请人工确认近 {span} 年文献占比是否达标", "参考文献")
    comp = rules["reference.composition"]["value"]
    if entries and comp.get("specified", True):
        jm = [e for e in entries if "[J]" in e or "[M]" in e]
        pct = round(100 * len(jm) / len(entries))
        low_pct = comp.get("journalOrBookMinPercent", 60)
        if pct >= low_pct:
            rep.ok("reference.composition",
                   f"期刊论文与学术著作 {len(jm)}/{len(entries)} 条（{pct}%）达标", "参考文献")
        else:
            rep.warn("reference.composition",
                     f"期刊论文与学术著作 {len(jm)}/{len(entries)} 条（{pct}%），低于 {low_pct}%", "参考文献")
        span = comp.get("recentWithinYears", 5)
        horizon = date.today().year - span + 1
        recent = [y for y in years if y >= horizon]
        rep.info("reference.composition",
                 f"近 {span} 年（≥{horizon}）文献约 {len(recent)}/{len(years)} 条，须人工确认",
                 "参考文献")
    elif entries:
        rep.info("reference.composition", comp.get("note", "本 profile 未规定文献构成比例，跳过"), "参考文献")


def check_ai_disclosure(doc: Docx, rep: Report, rules):
    want = rules["ai.disclosure"]["value"]
    if not want.get("required"):
        rep.info("ai.disclosure", want.get("note", "本 profile 不要求 AI 工具使用声明，跳过"), "")
        return
    title = want.get("sectionTitle", "人工智能工具使用声明")
    hits = [i for i, b in enumerate(doc.blocks)
            if b["kind"] == "p" and title in norm(b["text"])]
    if not hits:
        rep.fail("ai.disclosure", f"未找到「{title}」专节（本专业规范要求必设，放在正文末尾或参考文献之后）", "")
        return
    idx = hits[0]
    joined = "".join(b["text"] for b in doc.blocks[idx:idx + 40] if b["kind"] == "p")
    missing = [k for k in want.get("mustDeclare", []) if k not in joined]
    if missing:
        rep.warn("ai.disclosure", f"声明中未见：{'、'.join(missing)}（须人工确认）", title)
    else:
        rep.ok("ai.disclosure", f"已设「{title}」专节，含 {len(want.get('mustDeclare', []))} 项要素", title)
    if want.get("signed") and "签字" not in joined and "签名" not in joined:
        rep.warn("ai.disclosure", "声明末尾未见签字位置（本专业规范要求本人签字确认）", title)


def check_body_length(doc: Docx, rep: Report, rules):
    start = None
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] == "p" and doc.heading_level(blk["style"]) == 1:
            start = i
            break
    if start is None:
        rep.warn("body.length", "未找到一级标题，无法统计正文长度", "")
        return
    chunks = []
    for blk in doc.blocks[start:]:
        if blk["kind"] == "p" and blk["text"].strip() in ("参考文献", "致谢"):
            break
        if blk["kind"] == "p":
            chunks.append(blk["text"])
    count = len(CHINESE.findall("".join(chunks)))
    low = rules["body.length"]["value"]["minChineseChars"]
    if count >= low:
        rep.ok("body.length", f"正文中文约 {count} 字（含题注，不含页眉页脚）", "正文")
    else:
        rep.fail("body.length", f"正文中文约 {count} 字（应 ≥{low}）", "正文")


def check_binding_order(doc: Docx, rep: Report, rules):
    order = rules["binding.order"]["value"]["order"]
    matchers = {
        "封面": lambda t: ("毕业设计" in t or "毕业论文" in t) and "本科" in t,
        "原创性及知识产权声明": lambda t: "原创性" in t and "声明" in t,
        "声明": lambda t: "原创性" in t and "声明" in t,
        "摘要": lambda t: norm(t) == "摘要",
        "中文摘要": lambda t: norm(t) == "摘要",
        "Abstract": lambda t: t.strip() == "Abstract",
        "英文摘要": lambda t: t.strip() == "Abstract",
        "目录": lambda t: norm(t) == "目录",
        "参考文献": lambda t: t.strip() == "参考文献",
        "人工智能工具使用声明": lambda t: "人工智能工具使用声明" in norm(t),
        "致谢": lambda t: t.strip() == "致谢",
    }
    positions = {}
    for i, blk in enumerate(doc.blocks):
        if blk["kind"] != "p":
            continue
        text = blk["text"].strip()
        if not text:
            continue
        for name in order:
            if name not in positions and name in matchers and matchers[name](text):
                positions[name] = i
        if "正文" in order and "正文" not in positions and doc.heading_level(blk["style"]) == 1:
            positions["正文"] = i
        if len(positions) == len(order):
            break
    found = [n for n in order if n in positions]
    sequence = sorted(found, key=lambda n: positions[n])
    missing = [n for n in order if n not in positions]
    if sequence != found:
        rep.fail("binding.order", f"实际顺序 {' → '.join(sequence)}（应为 {' → '.join(order)}）", "")
    elif missing:
        rep.warn("binding.order",
                 f"已确认 {' → '.join(sequence)}；本稿未出现：{'、'.join(missing)}", "")
    else:
        rep.ok("binding.order", "装订顺序正确", "")


# ---------------------------------------------------------------------- main


CHECKS = [
    check_sections,
    check_page_numbers,
    check_headers,
    check_styles,
    check_heading_numbering,
    check_front_matter,
    check_toc,
    check_captions,
    check_tables,
    check_footnotes,
    check_references,
    check_ai_disclosure,
    check_body_length,
    check_binding_order,
]


def resolve_profile(data, name=None):
    profiles = data.get("profile", {})
    available = profiles.get("available", {})
    active = name or profiles.get("default")
    if active not in available:
        raise SystemExit(f"未知 profile：{active}（可选：{'、'.join(available)}）")
    return active, available[active]


def load_rules(profile=None):
    path = Path(__file__).resolve().parent.parent / "references" / "rules.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    active, info = resolve_profile(data, profile)
    table = {}
    for rule in data["rules"]:
        resolved = dict(rule)
        if "values" in rule:
            resolved["value"] = rule["values"][active]
        resolved["clause"] = rule.get("clauses", {}).get(active) or rule.get("clause", "")
        table[rule["id"]] = resolved
    return data, table, active, info


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="云南财经大学本科毕业论文 .docx 格式自检")
    parser.add_argument("path", nargs="?", help="待检查的 .docx")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--only", help="只输出指定规则，逗号分隔规则 id")
    parser.add_argument("--list-rules", action="store_true", help="列出全部规则 id")
    parser.add_argument("--list-profiles", action="store_true", help="列出可用规则口径（profile）")
    parser.add_argument("--profile", help="规则口径：zhonghua-cs-2023（默认）或 school-436")
    parser.add_argument("--heading-scheme",
                        help="正文标题字体口径（仅 school-436）：song=436 原文（小 4 加粗宋体）；black=黑体口径")
    args = parser.parse_args(argv)

    data, table, active, info = load_rules(args.profile)
    if args.list_profiles:
        default = data["profile"]["default"]
        for name, meta in data["profile"]["available"].items():
            mark = "*" if name == default else " "
            print(f"{mark} {name:<18} {meta['title']}  ({meta['issued']})")
        return 0
    if args.list_rules:
        print(f"# 规则口径 {active} —— {info['title']}")
        for rule in data["rules"]:
            resolved = table[rule["id"]]
            print(f"{rule['id']:<26} [{rule['severity']:<4}] {resolved['label']}  ({resolved['clause']})")
        return 0
    if not args.path:
        parser.error("缺少 .docx 路径")

    if args.heading_scheme:
        schemes = table["heading.font"]["value"].get("schemes", {})
        if args.heading_scheme not in schemes:
            parser.error(f"profile {active} 没有 {args.heading_scheme} 口径，可选：{'、'.join(schemes)}")
        table["heading.font"]["value"]["active"] = args.heading_scheme

    path = Path(args.path)
    doc = Docx(path)
    rep = Report(table)
    for check in CHECKS:
        try:
            check(doc, rep, table)
        except Exception as exc:  # 单项异常不影响其余检查
            rep.warn("check.error", f"检查 {check.__name__} 内部异常：{type(exc).__name__}: {exc}", "")

    findings = rep.findings
    if args.only:
        wanted = {x.strip() for x in args.only.split(",")}
        findings = [f for f in findings if f["id"] in wanted]
    fails = [f for f in findings if f["status"] == "FAIL"]
    warns = [f for f in findings if f["status"] == "WARN"]
    passes = [f for f in findings if f["status"] == "PASS"]
    infos = [f for f in findings if f["status"] == "INFO"]

    if args.json:
        print(json.dumps({
            "file": str(path),
            "profile": active,
            "source": info["source"],
            "summary": {"fail": len(fails), "warn": len(warns), "pass": len(passes), "info": len(infos)},
            "findings": findings,
        }, ensure_ascii=False, indent=2))
        return 1 if fails else 0

    print(f"检查文件：{path.name}")
    print(f"规则口径：{active} —— {info['title']}")
    print(f"规则来源：{info['source']}")
    print(f"声明：{data['meta']['status']}\n")
    if fails:
        print("=== 必须修改（FAIL）===")
        for f in fails:
            print(f"[{f['id']}] {f['message']}" + (f"  —— {f['where']}" if f["where"] else ""))
        print()
    if warns:
        print("=== 建议确认（WARN）===")
        for f in warns:
            print(f"[{f['id']}] {f['message']}" + (f"  —— {f['where']}" if f["where"] else ""))
        print()
    if passes:
        print("=== 通过（PASS）===")
        for f in passes:
            print(f"[{f['id']}] {f['message']}" + (f"  —— {f['where']}" if f["where"] else ""))
    if infos:
        print("\n=== 提示（INFO）===")
        for f in infos:
            print(f"[{f['id']}] {f['message']}" + (f"  —— {f['where']}" if f["where"] else ""))
    print(f"\n小计：FAIL {len(fails)}，WARN {len(warns)}，PASS {len(passes)}，INFO {len(infos)}")
    print("说明：本检查读取 OOXML 与样式定义，不渲染版面。页眉串节、目录点线、图表")
    print("      跨页断开等仍需导出 PDF 转图逐页核对；两者冲突时以渲染结果为准。")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
