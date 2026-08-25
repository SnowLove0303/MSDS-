# -*- coding: utf-8 -*-
"""批量下载 GB 标准全文（openstd.samr.gov.cn）
已验证流程: showGb(取验证页) -> gc(验证码图) -> verifyCode(OCR识别) -> viewGb(下载PDF)
已存在则跳过(断点续传)。"""
import sys, time, json, re
from pathlib import Path
import requests, ddddocr

BASE = "https://openstd.samr.gov.cn/bzgk/std"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Sec-Fetch-Dest": "document", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1", "Upgrade-Insecure-Requests": "1",
}

def fetch_std(hcno, out_path, ocr, max_try=8):
    """下载单个标准. 返回 ('ok'|'skip'|'fail', 详情)."""
    out = Path(out_path)
    if out.exists() and out.stat().st_size > 1000:
        return ("skip", {"size": out.stat().st_size})
    s = requests.Session()
    h = dict(HEADERS)
    h["Referer"] = f"{BASE}/showGb?type=download&hcno={hcno}"
    s.headers.update(h)
    # 1. 打开下载页
    s.get(f"{BASE}/showGb?type=download&hcno={hcno}", timeout=25)
    for i in range(max_try):
        try:
            r = s.get(f"{BASE}/gc?_{int(time.time()*1000)}", timeout=25)
            if r.status_code != 200 or len(r.content) < 500:
                continue
            code = ocr.classification(r.content).strip()
            if len(code) != 4:
                continue
            rv = s.post(f"{BASE}/verifyCode", data={"verifyCode": code}, timeout=25)
            if rv.text.strip() == "success":
                # 2. 延迟后下载(服务器生成文件)
                for k in range(10):
                    time.sleep(1.5)
                    rd = s.get(f"{BASE}/viewGb?hcno={hcno}", timeout=120)
                    if rd.status_code == 200 and len(rd.content) > 1000:
                        out.write_bytes(rd.content)
                        return ("ok", {"try": i+1, "code": code, "size": len(rd.content)})
                return ("fail", {"stage": "download_empty", "code": code})
        except Exception as exc:
            last_err = str(exc)
            time.sleep(1)
    return ("fail", {"stage": "captcha", "detail": last_err if "last_err" in dir() else "unknown"})

if __name__ == "__main__":
    hcno_json = sys.argv[1] if len(sys.argv) > 1 else "gb30000_hcno.json"
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("GB30000下载")
    out_dir.mkdir(parents=True, exist_ok=True)
    items = json.load(open(hcno_json, encoding="utf-8"))
    ocr = ddddocr.DdddOcr(show_ad=False)
    ok = skip = fail = 0
    fails = []
    for it in items:
        std = it["std"]; hcno = it["hcno"]
        fname = std.replace("/", "-").replace(" ", "")   # 文件名安全化: GB/T 16483 → GBT16483
        sub = out_dir / it.get("dir", "") if it.get("dir") else out_dir
        sub.mkdir(parents=True, exist_ok=True)
        dest = sub / f"{fname}.pdf"
        st, info = fetch_std(hcno, dest, ocr)
        if st == "ok": ok += 1; print(f"✅ {std}: {info['size']} 字节 (识别{info.get('code')}, 尝试{info.get('try')})")
        elif st == "skip": skip += 1; print(f"⏭ {std}: 已存在 {info['size']} 字节")
        else: fail += 1; fails.append(std); print(f"❌ {std}: {info}")
        time.sleep(0.5)
    print(f"\n完成: ✅{ok}  ⏭{skip}  ❌{fail}")
    if fails:
        print("失败项:", ", ".join(fails))
