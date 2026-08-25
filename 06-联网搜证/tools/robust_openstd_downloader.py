# -*- coding: utf-8 -*-
import requests, ddddocr, time
from pathlib import Path

# 经过验证的国家标准开放平台标准下载器
BASE = "https://openstd.samr.gov.cn/bzgk/std"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Upgrade-Insecure-Requests": "1"
}

ocr = ddddocr.DdddOcr(show_ad=False)

def download_openstd_pdf(hcno, out_path, max_attempts=8):
    out = Path(out_path)
    if out.exists() and out.stat().st_size > 5000:
        return ("skip", out.stat().st_size)
        
    s = requests.Session()
    s.headers.update(HEADERS)
    
    # 模拟访问详情页
    s.get(f"{BASE}/newGbInfo?hcno={hcno}", timeout=20)
    time.sleep(0.5)
    
    # 访问 download 页
    show_url = f"{BASE}/showGb?type=download&hcno={hcno}"
    s.headers["Referer"] = show_url
    s.get(show_url, timeout=20)
    
    for attempt in range(max_attempts):
        try:
            # 获取验证码
            r_gc = s.get(f"{BASE}/gc?_{int(time.time()*1000)}", timeout=15)
            if r_gc.status_code != 200 or len(r_gc.content) < 500:
                time.sleep(1)
                continue
                
            code = ocr.classification(r_gc.content).strip()
            if len(code) != 4:
                time.sleep(1)
                continue
                
            # 校验验证码
            r_verify = s.post(f"{BASE}/verifyCode", data={"verifyCode": code}, timeout=15)
            if r_verify.text.strip() == "success":
                # 轮询拉取生成的 PDF 文件
                for wait_step in range(15):
                    time.sleep(2)
                    r_pdf = s.get(f"{BASE}/viewGb?hcno={hcno}", timeout=120)
                    if r_pdf.status_code == 200 and len(r_pdf.content) > 5000:
                        out.parent.mkdir(parents=True, exist_ok=True)
                        out.write_bytes(r_pdf.content)
                        return ("ok", len(r_pdf.content))
                # 未就绪换一轮验证
                break
        except Exception as e:
            time.sleep(1.5)
    return ("fail", 0)

if __name__ == '__main__':
    st, size = download_openstd_pdf("4D487D68BF0BD87E68CE0EA68183DAD6", "GB15258-2009.pdf")
    print("Download result:", st, size)
