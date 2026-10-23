#!/usr/bin/env python3
"""Read one selected PowerPoint math object from an exported slide as Markdown."""

import base64
import io
import json
import re
import subprocess
import sys
import unicodedata
import zipfile

from lxml import etree


def flat(text):
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", "", text.translate(dict.fromkeys((0xFDD0, 0xFDEE, 0xFDEF))))


request = json.load(sys.stdin)
with zipfile.ZipFile(io.BytesIO(base64.b64decode(request["presentation"], validate=True))) as archive:
    slide = archive.getinfo("ppt/slides/slide1.xml")
    if slide.file_size > 5_000_000:
        raise ValueError("公式所在幻灯片的 XML 过大。")
    root = etree.fromstring(archive.read(slide))

namespaces = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
formulae = root.xpath(".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]", namespaces=namespaces)
selected = flat(request["formulaText"])
matches = [node for node in formulae if flat("".join(node.xpath(".//m:t/text()", namespaces=namespaces))) == selected]
if len(matches) != 1:
    raise ValueError("无法唯一匹配选中的原生公式，请只选中一条完整公式。")

stylesheet = etree.parse("/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl")
math = etree.XSLT(stylesheet)(matches[0]).getroot()
normal_mi = "{http://www.w3.org/1998/Math/MathML}mi"
for parent in math.iter():
    children = list(parent)
    for child in children:
        if child.getparent() is not parent or child.tag != normal_mi or child.get("mathvariant") != "normal":
            continue
        child.tag = "{http://www.w3.org/1998/Math/MathML}mtext"
        child.attrib.clear()
        while child.getnext() is not None and child.getnext().tag == normal_mi and child.getnext().get("mathvariant") == "normal":
            following = child.getnext()
            child.text = (child.text or "") + (following.text or "")
            parent.remove(following)
mathml = etree.tostring(math, encoding="utf-8")
markdown = subprocess.check_output(
    ["pandoc", "-f", "html", "-t", "markdown", "--wrap=none"],
    input=mathml,
    timeout=10,
).decode("utf-8").strip()
if not markdown.startswith("$") or "\ufdd0" in markdown:
    raise ValueError("公式没有成功转换为 Markdown。")
sys.stdout.write(markdown)
