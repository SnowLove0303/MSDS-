# -*- coding: utf-8 -*-
"""卫生标准网搜索 GBZ 标准, 找出详情页/下载链接."""
import sys, json, time
from playwright.sync_api import sync_playwright

CHROME = r'E:\MorenAnzhuangLujing\Chrome\Chrome\Application\chrome.exe'

def main():
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True, executable_path=CHROME, args=['--no-sandbox'])
        pg = b.new_page()
        # 先访问卫生标准网首页
        pg.goto('https://hbba.sacinfo.org.cn/', timeout=45000, wait_until='networkidle')
        pg.wait_for_timeout(1500)
        print('首页标题:', pg.title())
        print('首页URL:', pg.url)
        # 截图看布局
        pg.screenshot(path='gbz_home.png')
        # 尝试直接访问搜索页
        for url in [
            'https://hbba.sacinfo.org.cn/stdList?key=GBZ%202.1-2019',
            'https://hbba.sacinfo.org.cn/stdList?key=GBZ%202.1',
        ]:
            pg.goto(url, timeout=45000, wait_until='networkidle')
            pg.wait_for_timeout(2000)
            print(f'\n搜索页 {url}')
            print('  标题:', pg.title())
            body = pg.evaluate('() => document.body.innerText')
            print('  正文前800字:')
            print(body[:800])
            # 提取链接
            links = pg.evaluate('''() => Array.from(document.querySelectorAll('a')).map(a=>[a.innerText.trim(), a.href]).filter(x=>x[0].length>4)''')
            print('  链接:')
            for t, h in links[:30]:
                print('   ', t[:40], '→', h[:80])
            pg.screenshot(path='gbz_list.png')
        b.close()

if __name__ == '__main__':
    main()
