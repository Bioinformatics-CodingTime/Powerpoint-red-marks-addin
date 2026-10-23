import base64
import json
import subprocess
import sys
from pathlib import Path


root = Path(__file__).resolve().parents[1]
source = r"$$P_{\mathrm{emp}}=\frac{n_{\mathrm{null}\geq\mathrm{obs}}+1}{N_{\mathrm{null}}+1}$$"
raw = "\ufdd0𝑃\ufdeeemp\ufdef=\ufdd0\ufdd0𝑛\ufdeenull≥obs\ufdef+1\ufdee\ufdd0𝑁\ufdeenull\ufdef+1\ufdef"
pptx = subprocess.check_output(["pandoc", "-f", "markdown", "-t", "pptx"], input=source.encode())
request = {"presentation": base64.b64encode(pptx).decode(), "formulaText": raw}
result = subprocess.check_output([sys.executable, str(root / "extract_formula.py")], input=json.dumps(request).encode()).decode()
assert result == r"$$P_{\text{emp}} = \frac{n_{\text{null} \geq \text{obs}} + 1}{N_{\text{null}} + 1}$$"
request["formulaText"] = "x^2"
invalid = subprocess.run([sys.executable, str(root / "extract_formula.py")], input=json.dumps(request).encode(), capture_output=True)
assert invalid.returncode != 0
assert "无法唯一匹配" in invalid.stderr.decode()
print("原生公式提取检查通过。")
