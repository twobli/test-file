from html.parser import HTMLParser
from pathlib import Path
import re
p=Path(r"C:\Users\Windows\Desktop\生命医学原型图-HTML\基于基因的眼病_慢病诊断应用.html")
html=p.read_text(encoding='utf-8')
class P(HTMLParser):
    def __init__(self): super().__init__(); self.skip=0; self.cur=None; self.buf=[]
    def handle_starttag(self,tag,attrs):
        if tag in ('style','script'): self.skip+=1
        if tag in ('title','h1','h2','h3','h4','label','button','select','input','textarea','option','p','li','th','td') and self.skip==0:
            self.cur=tag; self.buf=[]
    def handle_endtag(self,tag):
        if tag in ('style','script') and self.skip: self.skip-=1
        if self.cur==tag:
            txt=' '.join(''.join(self.buf).split())
            if txt: print(f'{tag.upper()}: {txt}')
            self.cur=None; self.buf=[]
    def handle_data(self,data):
        if self.cur and self.skip==0: self.buf.append(data)
P().feed(html)
