# -*- coding: utf-8 -*-
"""
粉笔模考解析页 → 干净 A4 PDF（一条命令跑完，不再依赖对话）

用法:
    python fenbi_export.py "保存网页的文件夹" [-o 输出.pdf] [--no-pdf]

前提:
    1) 在浏览器里打开粉笔解析页，Ctrl+S → 保存类型选「网页，全部(*.htm;*.html)」
    2) 把保存出来的 .html 和 _files 文件夹放在同一个目录里
    3) 需要本机装有 Edge（默认路径见 EDGE_CANDIDATES）

它做四件事:
    抽取 130 道题的 HTML（题干/选项/答案/解析/考点）→ 抽取 5 份材料截图
    → 用自带 CSS 重排 → headless Edge 打印 A4 PDF
全程不联网、不调模型，跑一次只花几秒钟。
"""
import base64, json, os, re, shutil, socket, subprocess, sys, tempfile, time, urllib.request

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

CSS = """
@page { size: A4; margin: 12mm 11mm 12mm 11mm; }
* { box-sizing: border-box; }
html, body { margin:0; padding:0; background:#fff; }
body { font-family:"Microsoft YaHei","微软雅黑","SimSun",sans-serif;
       font-size:10pt; line-height:1.62; color:#000;
       -webkit-print-color-adjust:exact; print-color-adjust:exact; }
h1 { font-size:17pt; margin:0 0 2mm; }
.sub { font-size:9.5pt; color:#333; margin-bottom:4mm; }
h2.mod { font-size:13pt; margin:5mm 0 2mm; padding:1.6mm 3mm;
         border-left:4pt solid #000; background:#f2f2f2; break-after:avoid; }
h2.mod .modn { float:right; font-size:9pt; font-weight:normal; color:#555; }
.mat { border:.8pt solid #333; border-radius:2px; padding:3mm 3.5mm; margin-bottom:3mm;
       background:#fbfbfc; break-inside:avoid; }
.matlabel { font-family:"SimHei","黑体",sans-serif; font-size:9pt; letter-spacing:.2em;
            margin-bottom:1.6mm; }
.mat img { max-width:100%; height:auto; }
.card { border:.7pt solid #cfcfcf; border-radius:2px; padding:3mm 3.5mm;
        margin-bottom:4mm; break-inside:avoid; }
.title { display:flex; align-items:baseline; gap:3mm; border-bottom:.8pt solid #bbb;
         padding-bottom:1.4mm; margin-bottom:2mm; break-after:avoid; }
.title-index { font-family:"SimHei","黑体",sans-serif; font-size:10pt;
               background:#000; color:#fff; padding:.5mm 2mm; }
.title-type-name { font-size:9pt; color:#555; }
.title-right { display:none !important; }
.question-choice-container p { margin:1.2mm 0; text-align:justify; }
.question-choice-container img { max-width:100%; height:auto; }
ul.choice-radios { list-style:none; margin:1.6mm 0; padding:0; }
li.choice-radio { display:block; margin:1.1mm 0; break-inside:avoid; }
li.choice-radio .input-radio { display:inline-block !important; width:auto !important;
    height:auto !important; border:none !important; background:none !important;
    border-radius:0 !important; font-family:"SimHei","黑体",sans-serif;
    font-size:9.5pt !important; line-height:1.2 !important; margin-right:2mm;
    padding:0 !important; color:#000; }
li.choice-radio .input-radio > * { display:none !important; }
/* 你选的（对）*/
li.choice-radio:has(.input-radio.correct) .input-radio { color:#1a7f37; }
li.choice-radio:has(.input-radio.correct) .input-radio::after { content:" ✓"; font-size:8pt; }
li.choice-radio:has(.input-radio.correct) .input-text { font-weight:600; }
li.choice-radio:has(.input-radio.correct) { background:#f2fbf5; }
/* 你选的（错）*/
li.choice-radio:has(.input-radio.wrong) .input-radio { color:#c00; }
li.choice-radio:has(.input-radio.wrong) .input-radio::after { content:" ✗"; font-size:8pt; }
li.choice-radio:has(.input-radio.wrong) .input-text { font-weight:600; }
li.choice-radio:has(.input-radio.wrong) { background:#fff5f5; }
/* 正确答案（你没选它）*/
li.choice-radio:has(.input-radio.correctLost) .input-radio { color:#1a7f37; }
li.choice-radio:has(.input-radio.correctLost) .input-radio::after { content:" ✓"; font-size:8pt; }
.input-text { display:inline !important; margin:0; text-align:justify; }
.question-overall-container { display:flex; flex-wrap:wrap; gap:2mm 6mm;
      background:#f7f8fa; border-radius:2px; padding:2mm 3mm; margin:2.4mm 0;
      break-inside:avoid; }
.overall-item { display:flex; align-items:baseline; gap:1.4mm; font-size:8.6pt; }
.overall-item-value { font-family:"SimHei","黑体",sans-serif; }
.correct-answer { color:#1a7f37; } .correct-rate { color:#1a4fd6; }
.answer-time { color:#7a3ec8; } .error-prone { color:#c2410c; }
.your-answer { color:#c00; }
.overall-item-title, .overall-item-sub { color:#666; font-size:8.2pt; }
.result-common-container { height:auto !important; max-height:none !important;
                           overflow:visible !important; margin-top:1mm; }
.result-common-section { break-after:avoid; }
.solution-title-container { border-top:.7pt dashed #bbb; padding-top:1.6mm;
                            margin-top:1.6mm; break-after:avoid; }
.solution-title { font-family:"SimHei","黑体",sans-serif; font-size:9pt; letter-spacing:.16em; }
.content { display:block !important; height:auto !important; max-height:none !important;
           overflow:visible !important; }
.content p, .content div { margin:1mm 0; text-align:justify; }
.content img { max-width:100%; height:auto; }
.solution-keypoint-container { margin-top:1.4mm; }
.solution-keypoint-container * { font-size:8.6pt; }
"""

JS_DUMP = r"""
(() => {
  const out = [];
  document.querySelectorAll('app-ti').forEach(ti => {
    const box = ti.querySelector('.ti-container') || ti;
    const noEl = ti.querySelector('.title-index');
    const mine = [...ti.querySelectorAll('li.choice-radio')].map(li => {
      const r = li.querySelector('.input-radio');
      const c = r ? (r.className || '') : '';
      const k = r ? (r.textContent || '').trim().slice(0, 1) : '';
      return {k, c};
    }).filter(o => /(^|\s)(correct|wrong)(\s|$)/.test(o.c));
    out.push({
      no: noEl ? parseInt(noEl.textContent) : null,
      html: box.innerHTML,
      chosen: mine.length ? mine[0].k : null,
      imgs: [...ti.querySelectorAll('img')].map(i => i.getAttribute('src'))
    });
  });
  return JSON.stringify(out);
})()
"""

JS_MAT = r"""
(() => {
  const out = [];
  document.querySelectorAll('.materials-container').forEach(m => {
    const ti = m.closest('.ti');
    const t = ti ? ti.querySelector('.title-index') : null;
    const r = m.getBoundingClientRect();
    out.push({q: t ? parseInt(t.textContent) : null,
              x: r.left + window.scrollX, y: r.top + window.scrollY,
              w: r.width, h: r.height});
  });
  return JSON.stringify(out);
})()
"""


def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p


class CDP:
    def __init__(self, edge, profile):
        self.port = free_port()
        self.proc = subprocess.Popen(
            [edge, "--headless=new", "--disable-gpu", "--no-first-run",
             "--no-default-browser-check", "--allow-file-access-from-files",
             f"--remote-debugging-port={self.port}", f"--user-data-dir={profile}",
             "about:blank"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        base = f"http://127.0.0.1:{self.port}"
        for _ in range(200):
            try:
                urllib.request.urlopen(base + "/json/version", timeout=1).read()
                break
            except Exception:
                time.sleep(0.2)
        pages = json.loads(urllib.request.urlopen(base + "/json/list", timeout=3).read())
        page = next(t for t in pages if t.get("type") == "page")
        import websocket
        self.ws = websocket.create_connection(page["webSocketDebuggerUrl"], timeout=120)
        self.mid = 0

    def call(self, method, params=None):
        self.mid += 1
        self.ws.send(json.dumps({"id": self.mid, "method": method, "params": params or {}}))
        while True:
            d = json.loads(self.ws.recv())
            if d.get("id") == self.mid:
                if "error" in d:
                    raise RuntimeError(d["error"])
                return d.get("result", {})

    def js(self, expr):
        r = self.call("Runtime.evaluate",
                      {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if r.get("exceptionDetails"):
            raise RuntimeError(r["exceptionDetails"])
        return r["result"].get("value")

    def goto(self, url, wait=6):
        self.call("Page.enable")
        self.call("Runtime.enable")
        self.call("Page.navigate", {"url": url})
        time.sleep(wait)

    def shot_region(self, box, out_png, scale=2):
        r = self.call("Page.captureScreenshot", {
            "format": "png", "captureBeyondViewport": True,
            "clip": {"x": box["x"], "y": box["y"], "width": box["w"],
                     "height": box["h"], "scale": scale}})
        open(out_png, "wb").write(base64.b64decode(r["data"]))

    def pdf(self, out_pdf):
        r = self.call("Page.printToPDF", {
            "printBackground": True, "displayHeaderFooter": False,
            "paperWidth": 8.268, "paperHeight": 11.693,
            "marginTop": 0.472, "marginBottom": 0.472,
            "marginLeft": 0.433, "marginRight": 0.433,
            "preferCSSPageSize": False, "scale": 1.0})
        open(out_pdf, "wb").write(base64.b64decode(r["data"]))

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass
        self.proc.terminate()
        try:
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()


def clean(h):
    h = re.sub(r'<ul[^>]*class="title-right"[^>]*>.*?</ul>', "", h, flags=re.S)
    h = re.sub(r'<div[^>]*class="toggle-btn[^"]*"[^>]*>.*?</div>', "", h, flags=re.S)
    h = re.sub(r'\s_ngcontent-[^=]+="[^"]*"', "", h)
    h = re.sub(r'\s_nghost-[^=]+="[^"]*"', "", h)
    h = re.sub(r'\sng-reflect-[^=]+="[^"]*"', "", h)
    h = re.sub(r'style="[^"]*height:\s*0(px)?;[^"]*"', "", h)
    h = re.sub(r'style="[^"]*overflow:\s*hidden[^"]*"', "", h)
    return re.sub(r'<!---->', "", h)


MODS = [("政治理论", 1, 20), ("常识判断", 21, 35), ("言语理解与表达", 36, 65),
        ("数量关系", 66, 75), ("判断推理", 76, 110), ("资料分析", 111, 130)]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if not args:
        print(__doc__)
        return 1
    folder = os.path.abspath(args[0])
    src = None
    if os.path.isfile(folder) and folder.lower().endswith((".html", ".htm")):
        src = folder
        folder = os.path.dirname(folder)
    out_pdf = os.path.join(folder, "模考_题目+我的答案+完整解析.pdf")
    for i, a in enumerate(sys.argv):
        if a == "-o" and i + 1 < len(sys.argv):
            out_pdf = os.path.abspath(sys.argv[i + 1])
    if src is None:
        htmls = sorted([f for f in os.listdir(folder)
                        if f.lower().endswith((".html", ".htm")) and not f.startswith("_")])
        if not htmls:
            print("!! 这个文件夹里没找到 .html。")
            print("   请确认为「网页，全部」保存，.html 和 _files 文件夹要在同一层。")
            return 1
        if len(htmls) > 1:
            print("   发现多个 html，使用第一个:", htmls[0])
        src = os.path.join(folder, htmls[0])
    edge = next((e for e in EDGE_CANDIDATES if os.path.exists(e)), None)
    if not edge:
        print("!! 没找到 Edge/Chrome")
        return 1

    # 所有中间产物都放系统临时目录，跑完自动删除，不在你桌面上留垃圾
    work = tempfile.mkdtemp(prefix="fenbi_export_")

    print(f"[1/4] 打开 {os.path.basename(src)}")
    b = CDP(edge, os.path.join(work, "_edge"))
    try:
        b.goto("file:///" + src.replace("\\", "/"))
        raw = b.js(JS_DUMP)
        arr = json.loads(raw)
        n_ok = sum(1 for q in arr if q["chosen"])
        print(f"      题目 {len(arr)} 道，读到“我选的项” {n_ok} 道")

        print("[2/4] 抓材料块")
        b.call("Emulation.setDeviceMetricsOverride",
               {"width": 1100, "height": 1600, "deviceScaleFactor": 2, "mobile": False})
        time.sleep(1.2)
        boxes = json.loads(b.js(JS_MAT))
        mats = {}
        for it in boxes:
            if it["h"] < 10:
                continue
            p = os.path.join(work, f"mat_{it['q']}.png")
            b.shot_region(it, p)
            mats[it["q"]] = p
        print(f"      材料 {len(mats)} 份")

        print("[3/4] 重排 HTML")
        parts = []
        for name, a, z in MODS:
            parts.append(f'<h2 class="mod">{name}<span class="modn">{z - a + 1} 题</span></h2>')
            for q in sorted([x for x in arr if x["no"] and a <= x["no"] <= z],
                            key=lambda x: x["no"]):
                if q["no"] in mats:
                    parts.append('<section class="mat"><div class="matlabel">材&nbsp;&nbsp;料</div>'
                                 f'<img src="mat_{q["no"]}.png"></section>')
                parts.append(f'<section class="q"><div class="card">{clean(q["html"])}</div></section>')
        head = ("<!doctype html><html lang=zh-CN><head><meta charset=utf-8>"
                "<title>模考 · 题目+我的答案+完整解析</title><style>" + CSS + "</style></head><body>"
                f'<h1>行测模考 · 共 {len(arr)} 题</h1>'
                '<div class="sub">题目 + 我选的项（✓ 选中） + 正确答案 + 完整解析</div>')
        page = head + "\n".join(parts) + "</body></html>"
        out_html = os.path.join(work, "_export.html")
        open(out_html, "w", encoding="utf-8").write(page)
        for _ in range(20):
            if os.path.exists(out_html) and os.path.getsize(out_html) > 1000:
                break
            time.sleep(0.2)

        if "--no-pdf" not in sys.argv:
            print("[4/4] 打印 PDF")
            b.goto("file:///" + out_html.replace("\\", "/"), wait=5)
            b.pdf(out_pdf)
            print("      ->", out_pdf, f"({os.path.getsize(out_pdf)/1024/1024:.1f} MB)")
        if "--tsv" in sys.argv:
            tsv = os.path.join(folder, "作答对照.tsv")
            with open(tsv, "w", encoding="utf-8") as f:
                f.write("题号\t我选的\n")
                for q in sorted(arr, key=lambda x: x["no"] or 0):
                    f.write(f'{q["no"]}\t{q["chosen"] or ""}\n')
            print("      对照表 ->", tsv)
    finally:
        b.close()
        shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
