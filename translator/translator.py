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
  {"cmd":"ping"}                        -> {"pong":true}
  {"cmd":"installed_pairs"}             -> {"pairs":[["en","pt"],...]}
  {"cmd":"install","from":"en","to":"pt"} -> {"ok":true}
  {"cmd":"available_packages"}          -> {"packages":[...]}
"""

import sys, json, threading

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
