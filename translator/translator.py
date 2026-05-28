#!/usr/bin/env python3
"""
DocTranslate Translation Microservice
──────────────────────────────────────
Runs as a background process launched by the C# WPF app.
Communicates via stdin/stdout using JSON lines.

Protocol:
  C# sends:     {"id":"1","text":"Hello","from":"en","to":"pt"}
  Python sends: {"id":"1","result":"Olá"}
  On error:     {"id":"1","error":"message"}

Commands:
  {"cmd":"ping"}                                      -> {"pong":true}
  {"cmd":"installed_pairs"}                           -> {"pairs":[["en","pt"],...]}
  {"cmd":"install","from":"en","to":"pt"}             -> {"ok":true}
  {"cmd":"available_packages"}                        -> {"packages":[...]}
  {"cmd":"ocr","id":"ocr_1","image_b64":"...","from_langs":["en"],"to":"pt"}
                                                      -> {"id":"ocr_1","text":"...","translated":"..."}
  {"cmd":"ocr_pdf_page","id":"ocr_2","path":"...","page":0,"from_langs":["en"],"to":"pt"}
                                                      -> {"id":"ocr_2","text":"...","translated":"..."}
"""

import sys, json, threading, base64

try:
    import argostranslate.package, argostranslate.translate
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,"-m","pip","install","argostranslate","-q"])
    import argostranslate.package, argostranslate.translate

_lang_cache = None
_cache_lock = threading.Lock()

def _get_lang_map():
    global _lang_cache
    with _cache_lock:
        if _lang_cache is None:
            langs = argostranslate.translate.get_installed_languages()
            _lang_cache = {l.code: l for l in langs}
        return _lang_cache

def _invalidate_cache():
    global _lang_cache
    with _cache_lock:
        _lang_cache = None

def translate(text, from_code, to_code):
    text = text.strip()
    if not text: return text
    lang_map = _get_lang_map()
    src = lang_map.get(from_code)
    tgt = lang_map.get(to_code)
    if not src or not tgt: return text
    t = src.get_translation(tgt)
    if t: return t.translate(text)
    if from_code != "en" and to_code != "en":
        en = lang_map.get("en")
        if en:
            to_en    = src.get_translation(en)
            en_to_tg = en.get_translation(tgt)
            if to_en and en_to_tg:
                return en_to_tg.translate(to_en.translate(text))
    return text

def get_installed_pairs():
    try:
        return [[p.from_code, p.to_code]
                for p in argostranslate.package.get_installed_packages()]
    except: return []

def install_pair(from_code, to_code):
    try:
        argostranslate.package.update_package_index()
        available = argostranslate.package.get_available_packages()
        pkg = next((p for p in available
                    if p.from_code == from_code and p.to_code == to_code), None)
        if not pkg: return False
        argostranslate.package.install_from_path(pkg.download())
        _invalidate_cache()
        return True
    except: return False

def get_available_packages():
    try:
        argostranslate.package.update_package_index()
        return [{"from":p.from_code,"to":p.to_code,
                 "from_name":p.from_name,"to_name":p.to_name}
                for p in argostranslate.package.get_available_packages()]
    except: return []

# ── OCR ───────────────────────────────────────────────────────────────────────

_ocr_reader    = None
_ocr_reader_langs: list = []
_ocr_lock      = threading.Lock()
_ocr_available: bool | None = None   # None=untested, True=ok, False=missing

def _get_ocr_reader(langs):
    """Lazy-init EasyOCR reader. Reuses existing reader if langs match."""
    global _ocr_reader, _ocr_reader_langs, _ocr_available
    if _ocr_available is False:
        return None
    with _ocr_lock:
        if _ocr_reader is None or sorted(langs) != sorted(_ocr_reader_langs):
            try:
                import easyocr
                _ocr_reader = easyocr.Reader(langs, gpu=False, verbose=False)
                _ocr_reader_langs = langs[:]
                _ocr_available = True
            except ImportError:
                _ocr_available = False
                return None
        return _ocr_reader

def _ocr_bytes(image_bytes: bytes, langs: list) -> str:
    reader = _get_ocr_reader(langs)
    if reader is None:
        return ""
    results = reader.readtext(image_bytes, detail=0, paragraph=True)
    return " ".join(r for r in results if r).strip()

def _rasterize_pdf_page(pdf_path: str, page_index: int, dpi: int = 150) -> bytes:
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise RuntimeError("pymupdf not installed — run: pip install pymupdf")
    doc  = fitz.open(pdf_path)
    page = doc[page_index]
    mat  = fitz.Matrix(dpi / 72, dpi / 72)
    pix  = page.get_pixmap(matrix=mat)
    return pix.tobytes("png")

def handle_ocr(req_id: str, image_bytes: bytes, from_langs: list, to_code: str):
    try:
        original   = _ocr_bytes(image_bytes, from_langs)
        from_code  = from_langs[0] if from_langs else "en"
        translated = translate(original, from_code, to_code) if original else ""
        print(json.dumps({"id": req_id, "text": original, "translated": translated}),
              flush=True)
    except Exception as e:
        print(json.dumps({"id": req_id, "text": "", "translated": "", "error": str(e)}),
              flush=True)

def handle_ocr_pdf_page(req_id: str, pdf_path: str, page_index: int,
                         from_langs: list, to_code: str):
    try:
        image_bytes = _rasterize_pdf_page(pdf_path, page_index)
        handle_ocr(req_id, image_bytes, from_langs, to_code)
    except Exception as e:
        print(json.dumps({"id": req_id, "text": "", "translated": "", "error": str(e)}),
              flush=True)

# ── Main loop ─────────────────────────────────────────────────────────────────

def main():
    print(json.dumps({"ready": True}), flush=True)
    for raw_line in sys.stdin:
        raw_line = raw_line.strip()
        if not raw_line: continue
        try:
            msg = json.loads(raw_line)
        except json.JSONDecodeError as e:
            print(json.dumps({"error": f"JSON parse error: {e}"}), flush=True)
            continue
        try:
            if "cmd" in msg:
                cmd = msg["cmd"]
                if cmd == "ping":
                    print(json.dumps({"pong": True}), flush=True)

                elif cmd == "installed_pairs":
                    print(json.dumps({"pairs": get_installed_pairs()}), flush=True)

                elif cmd == "available_packages":
                    print(json.dumps({"packages": get_available_packages()}), flush=True)

                elif cmd == "install":
                    ok = install_pair(msg.get("from",""), msg.get("to",""))
                    print(json.dumps({"ok": ok}), flush=True)

                elif cmd == "ocr":
                    req_id     = msg.get("id", "")
                    b64        = msg.get("image_b64", "")
                    from_langs = msg.get("from_langs", ["en"])
                    to_code    = msg.get("to", "pt")
                    image_bytes = base64.b64decode(b64)
                    # Run in a thread so the stdin loop stays responsive
                    threading.Thread(
                        target=handle_ocr,
                        args=(req_id, image_bytes, from_langs, to_code),
                        daemon=True
                    ).start()

                elif cmd == "ocr_pdf_page":
                    req_id     = msg.get("id", "")
                    pdf_path   = msg.get("path", "")
                    page_index = msg.get("page", 0)
                    from_langs = msg.get("from_langs", ["en"])
                    to_code    = msg.get("to", "pt")
                    threading.Thread(
                        target=handle_ocr_pdf_page,
                        args=(req_id, pdf_path, page_index, from_langs, to_code),
                        daemon=True
                    ).start()

                else:
                    print(json.dumps({"error": f"Unknown cmd: {cmd}"}), flush=True)

            elif "text" in msg:
                req_id    = msg.get("id","")
                result    = translate(msg["text"], msg.get("from","en"), msg.get("to","pt"))
                print(json.dumps({"id": req_id, "result": result}), flush=True)

            else:
                print(json.dumps({"error": "Unknown message format"}), flush=True)

        except Exception as e:
            req_id = msg.get("id","") if isinstance(msg, dict) else ""
            print(json.dumps({"id": req_id, "error": str(e)}), flush=True)

if __name__ == "__main__":
    main()
