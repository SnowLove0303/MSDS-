# -*- coding: utf-8 -*-
"""补下脚本: 对 openstd 生成慢/限流的标准做多轮强化重试.
用法: python gb_retry_download.py <清单.json> <输出目录> [轮数]
已存在(>1000字节)则跳过."""
import sys, time, json
from pathlib import Path
import requests, ddddocr

BASE = "https://openstd.samr.gov.cn/bzgk/std"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

def fetch_std(hcno, out_path, ocr, rounds=4):
    """多轮重试下载. 每轮: showGb -> 验证码 -> viewGb(等 15 次 x 2s)."""
    out = Path(out_path)
    if out.exists() and out.stat().st_size > 1000:
        return ("skip", {"size": out.stat().st_size})
    for rnd in range(rounds):
        s = requests.Session()
        h = dict(HEADERS)
        h["Referer"] = f"{BASE}/showGb?type=download&hcno={hcno}"   # 全局头关键
        s.headers.update(h)
        try:
            s.get(f"{BASE}/showGb?type=download&hcno={hcno}", timeout=25)
        except Exception:
            time.sleep(3); continue
        for i in range(6):
            try:
                r = s.get(f"{BASE}/gc?_{int(time.time()*1000)}", timeout=25)
                if r.status_code != 200 or len(r.content) < 500:
                    continue
                code = ocr.classification(r.content).strip()
                if len(code) != 4:
                    continue
                rv = s.post(f"{BASE}/verifyCode", data={"verifyCode": code}, timeout=25)
                if rv.text.strip() == "success":
                    for k in range(15):
                        time.sleep(2)
                        rd = s.get(f"{BASE}/viewGb?hcno={hcno}", timeout=180)
                        if rd.status_code == 200 and len(rd.content) > 1000:
                            out.write_bytes(rd.content)
                            return ("ok", {"round": rnd+1, "code": code, "size": len(rd.content)})
                    break     # 验证码成功但文件未生成 -> 换一轮
            except Exception:
                time.sleep(2)
    return ("fail", {"stage": "empty_after_retry"})

if __name__ == "__main__":
    hcno_json = sys.argv[1]
    out_dir = Path(sys.argv[2])
    rounds = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    items = json.load(open(hcno_json, encoding="utf-8"))
    ocr = ddddocr.DdddOcr(show_ad=False)
    ok = skip = fail = 0
    for it in items:
        std = it["std"]
        fname = std.replace("/", "-").replace(" ", "")
        sub = out_dir / it.get("dir", "") if it.get("dir") else out_dir
        sub.mkdir(parents=True, exist_ok=True)
        dest = sub / f"{fname}.pdf"
        st, info = fetch_std(it["hcno"], dest, ocr, rounds)
        if st == "ok": ok += 1; print(f"✅ {std}: {info['size']} 字节 (第{info['round']}轮, {info.get('code')})")
        elif st == "skip": skip += 1; print(f"⏭ {std}: 已存在 {info['size']} 字节")
        else: fail += 1; print(f"❌ {std}: {info}")
        time.sleep(1)
    print(f"\n完成: ✅{ok}  ⏭{skip}  ❌{fail}")
