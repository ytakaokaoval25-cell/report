#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""実験報告書 v7: SEM全30画像埋め込み + 第4章大幅改訂"""
import io, math, copy
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['Hiragino Sans', 'Yu Gothic', 'Arial Unicode MS', 'DejaVu Sans']
rcParams['axes.unicode_minus'] = False

try:
    from scipy.interpolate import PchipInterpolator
    _SCIPY = True
except ImportError:
    _SCIPY = False

from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_PATH = "/Users/takaoka/report/report_final.docx"

# ============================================================
# 実測データ（CSV より）
# ============================================================
SS400_ENG = [
    (0.0,0.0),(0.00248,14.17),(0.00498,30.91),(0.00746,52.17),
    (0.00994,76.64),(0.01242,101.76),(0.01554,131.39),(0.01864,160.37),
    (0.02174,193.21),(0.02486,226.70),(0.02796,260.84),(0.03107,294.97),
    (0.03418,328.46),(0.03728,368.40),(0.03977,392.87),(0.04194,407.04),
    (0.04412,415.73),  # 上降伏点
    (0.04661,354.23),  # 降伏低下
    (0.04909,341.35),  # 下降伏点
    (0.05157,344.57),(0.05437,352.94),(0.05903,365.82),(0.06525,379.99),
    (0.07457,392.87),(0.08699,407.04),(0.09942,418.63),(0.11185,428.29),
    (0.12428,436.66),(0.13982,443.10),(0.15534,448.26),(0.17398,452.12),
    (0.19263,455.99),(0.21127,458.56),(0.22681,460.49),(0.24233,461.78),
    (0.25477,462.43),(0.26253,462.75),(0.26719,463.07),  # σB ピーク
    (0.27185,454.05),(0.27807,425.07),
    (0.28584,312.36),  # 破断直前
]
FC250_ENG = [
    (0.0,0.0),(0.0026,32.0),(0.004,56.0),(0.0054,118.0),
    (0.0075,236.0),(0.0086,298.0),(0.0098,347.0),(0.0108,396.14),
]
A7075_ENG = [
    (0.0,0.0),(0.008,110.0),(0.018,260.0),(0.030,520.0),
    (0.042,800.0),(0.053,980.0),(0.063,1060.0),(0.071,1103.19),
    (0.07686,1040.0),
]
A7075_TRUE = [
    (0.0,0.0),(0.007968,110.88),(0.017840,264.68),(0.029559,535.60),
    (0.041142,833.60),(0.051643,1031.94),(0.061095,1126.78),
    (0.068593,1181.52),(0.074049,1119.93),
]
FC250_TRUE = [
    (0.0,0.0),(0.002597,32.08),(0.003992,56.22),(0.005385,118.64),
    (0.007472,237.77),(0.008563,300.56),(0.009752,350.40),(0.010742,400.42),
]

# 実測キー値
KV = {
    'SS400': dict(sy=415.73, sL=341.35, sB=463.07, Pmax=71.9,
                  phi=22.20, delta=28.58, E_lit=206, sy_lit=245, sB_lit=400),
    'FC250': dict(sy=None, sB=396.14, Pmax=61.2,
                  phi=4.61, delta=4.84, E_lit=130, sy_lit=None, sB_lit=250),
    'A7075': dict(sy=None, sB=1103.19, Pmax=170.0,
                  phi=7.14, delta=7.69, E_lit=71.7, sy_lit=503, sB_lit=572),
}

MAT = {
    'SS400':     dict(d1=14.060,d2=14.053,d3=14.068,l0=82.18,A0=155.168,lf=105.67,
                     E_lit=206,sy_lit=245,sB_lit=400,sf_lit=320),
    'FC250':     dict(d1=14.060,d2=13.992,d3=14.022,l0=62.22,A0=154.491,lf=65.23,
                     E_lit=130,sy_lit=None,sB_lit=250,sf_lit=None),
    'A7075':     dict(d1=13.982,d2=14.017,d3=14.023,l0=62.45,A0=154.099,lf=67.25,
                     E_lit=71.7,sy_lit=503,sB_lit=572,sf_lit=450),
    'SS400_add': dict(d1=13.980,d2=13.981,d3=14.070,l0=82.30,A0=154.165,lf=110.18,
                     E_lit=206,sy_lit=245,sB_lit=400,sf_lit=320),
}
CHARPY = {
    'SS400': dict(l=54.98,b=10.00,h1=10.00,h2=8.53,A=100.0),
    'FC250': dict(l=54.90,b=10.00,h1=9.97, h2=8.25,A=99.7),
    'A7075': dict(l=55.00,b=9.90, h1=10.00,h2=8.30,A=99.0),
}
W_kg=25.093; g_acc=9.80665; W_N=W_kg*g_acc; r_m=0.6513; Wr=W_N*r_m
alpha=146.5; cos_a=math.cos(math.radians(alpha))

def elong(key):
    d=MAT[key]; return (d['lf']-d['l0'])/d['l0']*100

# ============================================================
# フォント・段落ヘルパー
# ============================================================
def _set_run_font(run, size_pt=10.5, bold=False, italic=False, jp_font='ＭＳ 明朝'):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Times New Roman'
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'),    'Times New Roman')
    rFonts.set(qn('w:hAnsi'),    'Times New Roman')
    rFonts.set(qn('w:eastAsia'), jp_font)
    rFonts.set(qn('w:cs'),       'Times New Roman')  # Windows互換
    for e in rPr.findall(qn('w:rFonts')): rPr.remove(e)
    rPr.insert(0, rFonts)

def _set_spacing(para, pt=18):
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing'); pPr.append(sp)
    sp.set(qn('w:line'), str(int(Pt(pt).pt*20)))
    sp.set(qn('w:lineRule'), 'exact')

def para(doc, text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         size=10.5, bold=False, italic=False, sb=0, sa=3, indent_mm=4, line_pt=18):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    p.paragraph_format.first_line_indent = Mm(indent_mm)
    _set_spacing(p, line_pt)
    if text:
        r = p.add_run(text)
        _set_run_font(r, size, bold, italic)
    return p

def chapter_title(doc, text):
    return para(doc, text, size=15, bold=True, sb=14, sa=6, indent_mm=0)
def h1(doc, text):
    return para(doc, text, size=13, bold=True, sb=10, sa=4, indent_mm=0)
def h2(doc, text):
    return para(doc, text, size=11, bold=True, sb=6,  sa=3, indent_mm=0)
def h3(doc, text):
    return para(doc, text, size=10.5, bold=True, sb=5, sa=2, indent_mm=0)

def equation(doc, eq_text, eq_num):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Mm(8)
    _set_spacing(p, 16)
    r1 = p.add_run(eq_text); _set_run_font(r1, 10.5, italic=True)
    r2 = p.add_run(f'\t({eq_num})'); _set_run_font(r2, 10.5)
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'),'right'); tab.set(qn('w:pos'),str(int(Mm(155).pt*20)))
    tabs.append(tab); pPr.append(tabs)
    return p

def fig_caption(doc, fn, text):
    return para(doc, f'Fig. {fn}　{text}',
                align=WD_ALIGN_PARAGRAPH.CENTER, size=10, sb=1, sa=6, indent_mm=0)
def tbl_caption(doc, tn, text):
    return para(doc, f'Table {tn}　{text}',
                align=WD_ALIGN_PARAGRAPH.CENTER, size=10, sb=6, sa=2, indent_mm=0)

def set_margins(doc):
    for sec in doc.sections:
        sec.page_height=Mm(297); sec.page_width=Mm(210)
        sec.left_margin=Mm(25);  sec.right_margin=Mm(15)
        sec.top_margin=Mm(25);   sec.bottom_margin=Mm(20)

def fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0); plt.close(fig); return buf

def insert_figure(doc, buf, width_cm=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(buf, width=Cm(width_cm))
    return p

# ============================================================
# Unicode 添字・上付きを Word 書式に変換するポストプロセッサ
# ============================================================
def fix_subscripts_in_doc(doc):
    SUB_MAP = {'0':'0','1':'1','2':'2','3':'3','4':'4',
               '5':'5','6':'6','7':'7','8':'8','9':'9',
               '₀':'0','₁':'1','₂':'2','₃':'3','₄':'4',
               '₅':'5','₆':'6','₇':'7','₈':'8','₉':'9',
               'ₐ':'a','ᵢ':'i','ⱼ':'j','ₚ':'p','ₜ':'t',
               'ₑ':'e','ₒ':'o','ₙ':'n','ₛ':'s','ᵤ':'u'}
    SUP_MAP = {'²':'2','³':'3','⁰':'0','⁴':'4','⁵':'5',
               '⁶':'6','⁷':'7','⁸':'8','⁹':'9',
               'ⁿ':'n','ᵖ':'p','ᵃ':'a','ᵇ':'b','ᶜ':'c','ᵉ':'e'}
    # subscript digit chars in range 0x2080-0x2089 already mapped above
    ALL = {k:('subscript',v) for k,v in SUB_MAP.items()}
    ALL.update({k:('superscript',v) for k,v in SUP_MAP.items()})

    def split_text(text):
        segs = []
        cur=''; ctype='normal'
        for c in text:
            if c in ALL:
                if cur: segs.append((cur,ctype))
                t,ch = ALL[c]; segs.append((ch,t))
                cur=''; ctype='normal'
            else:
                if ctype!='normal' and cur:
                    segs.append((cur,ctype)); cur=''; ctype='normal'
                cur+=c
        if cur: segs.append((cur,ctype))
        # merge consecutive same type
        merged=[]
        for s in segs:
            if merged and merged[-1][1]==s[1]:
                merged[-1]=(merged[-1][0]+s[0],s[1])
            else:
                merged.append(list(s))
        return merged

    def process_run(run_xml, parent):
        t_el = run_xml.find(qn('w:t'))
        if t_el is None: return
        text = t_el.text or ''
        if not any(c in ALL for c in text): return
        segs = split_text(text)
        if len(segs)==1 and segs[0][1]=='normal': return
        idx = list(parent).index(run_xml)
        new_rs=[]
        for seg_text, seg_type in segs:
            nr = copy.deepcopy(run_xml)
            for t in nr.findall(qn('w:t')): nr.remove(t)
            t = OxmlElement('w:t'); t.text = seg_text
            if seg_text != seg_text.strip() or seg_text==' ':
                t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            nr.append(t)
            if seg_type != 'normal':
                rPr = nr.find(qn('w:rPr'))
                if rPr is None:
                    rPr=OxmlElement('w:rPr'); nr.insert(0,rPr)
                for va in rPr.findall(qn('w:vertAlign')): rPr.remove(va)
                va=OxmlElement('w:vertAlign')
                va.set(qn('w:val'),seg_type); rPr.append(va)
                for tag in (qn('w:sz'),qn('w:szCs')):
                    el=rPr.find(tag)
                    if el is not None:
                        v=int(el.get(qn('w:val'),'21'))
                        el.set(qn('w:val'),str(max(14,int(v*0.72))))
            new_rs.append(nr)
        for i,nr in enumerate(new_rs): parent.insert(idx+i,nr)
        parent.remove(run_xml)

    def fix_para(p_xml):
        for r in list(p_xml.findall(qn('w:r'))):
            try: process_run(r, p_xml)
            except: pass

    for p in doc.paragraphs: fix_para(p._p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs: fix_para(p._p)

# ============================================================
# グラフ生成（実測データ使用）
# ============================================================
def _smooth(e_arr, s_arr, n=300):
    if _SCIPY and len(e_arr)>3:
        cs = PchipInterpolator(e_arr, s_arr)
        ex = np.linspace(e_arr[0], e_arr[-1], n)
        return ex, cs(ex)
    ex = np.linspace(e_arr[0], e_arr[-1], n)
    return ex, np.interp(ex, e_arr, s_arr)

def plot_overview():
    fig, ax = plt.subplots(figsize=(10,6))
    # SS400
    e,s = zip(*SS400_ENG)
    ax.plot(np.array(e)*100, s, color='steelblue', lw=2, label='SS400（構造用鋼）')
    # A7075
    e,s = zip(*A7075_ENG); e,s = np.array(e),np.array(s)
    ex,sx = _smooth(e,s)
    ax.plot(ex*100, sx, color='darkorange', ls='--', lw=2, label='A7075（Al合金）')
    # FC250
    e,s = zip(*FC250_ENG)
    ax.plot(np.array(e)*100, s, color='forestgreen', ls=':', lw=2, label='FC250（ねずみ鋳鉄）')
    ax.set_xlabel('公称ひずみ e [%]', fontsize=12)
    ax.set_ylabel('公称応力 sigma [MPa]', fontsize=12)
    ax.set_title('3材料 公称応力-公称ひずみ線図（実測データ）', fontsize=13)
    ax.legend(fontsize=11); ax.set_xlim(0); ax.set_ylim(0); ax.grid(True,ls='--',alpha=0.5)
    ax.text(0.35,0.97,'※FC250は参考データ（実測 delta=4.84%，CSV上は1.08%まで）',
            transform=ax.transAxes,fontsize=8,color='gray',va='top')
    fig.tight_layout(); return fig_to_buf(fig)

def plot_ss400():
    fig,(ax1,ax2)=plt.subplots(1,2,figsize=(13,5.5))
    e_arr=np.array([p[0] for p in SS400_ENG])
    s_arr=np.array([p[1] for p in SS400_ENG])
    ax1.plot(e_arr*100,s_arr,'steelblue',lw=2)
    ax1.plot(e_arr*100,s_arr,'o',color='steelblue',ms=3,alpha=0.5)
    # key annotations
    ax1.annotate(f'上降伏点 {KV["SS400"]["sy"]:.1f} MPa',
                 xy=(SS400_ENG[16][0]*100,KV['SS400']['sy']),
                 xytext=(SS400_ENG[16][0]*100+2,KV['SS400']['sy']*1.10),
                 fontsize=8,color='navy',arrowprops=dict(arrowstyle='->',color='navy',lw=1))
    ax1.annotate(f'下降伏点 {KV["SS400"]["sL"]:.1f} MPa',
                 xy=(SS400_ENG[18][0]*100,KV['SS400']['sL']),
                 xytext=(SS400_ENG[18][0]*100+2,KV['SS400']['sL']*0.88),
                 fontsize=8,color='navy',arrowprops=dict(arrowstyle='->',color='navy',lw=1))
    ax1.annotate(f'sigma_B {KV["SS400"]["sB"]:.1f} MPa',
                 xy=(SS400_ENG[37][0]*100,KV['SS400']['sB']),
                 xytext=(SS400_ENG[37][0]*100-6,KV['SS400']['sB']*1.06),
                 fontsize=8,color='darkred',arrowprops=dict(arrowstyle='->',color='darkred',lw=1))
    ax1.axvline(KV['SS400']['delta'],color='red',ls='--',lw=1,
                label=f'破断 delta={KV["SS400"]["delta"]:.2f}%')
    ax1.set_xlabel('公称ひずみ e [%]',fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]',fontsize=11)
    ax1.set_title('SS400\n公称応力-公称ひずみ線図（実測）',fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True,ls='--',alpha=0.5)
    # true stress (up to peak pt38, index 0..37)
    e_u=e_arr[:38]; s_u=s_arr[:38]
    et=np.log(1+e_u); st=s_u*(1+e_u)
    ax2.plot(et*100,st,'darkorange',lw=2)
    ax2.plot(et*100,st,'o',color='darkorange',ms=3,alpha=0.5)
    ax2.set_xlabel('対数ひずみ epsilon [%]',fontsize=11)
    ax2.set_ylabel('真応力 σt [MPa]',fontsize=11)
    ax2.set_title('SS400\n真応力-対数ひずみ線図（均一変形域，実測）',fontsize=11)
    ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True,ls='--',alpha=0.5)
    ax2.text(0.02,0.03,'注：最大荷重点以降は体積一定仮定不成立のため除外',
             transform=ax2.transAxes,fontsize=8,color='gray')
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def plot_a7075():
    fig,(ax1,ax2)=plt.subplots(1,2,figsize=(13,5.5))
    e,s=zip(*A7075_ENG); e,s=np.array(e),np.array(s)
    ex,sx=_smooth(e,s)
    ax1.plot(ex*100,sx,'steelblue',lw=2)
    ax1.plot(e*100,s,'o',color='steelblue',ms=5,label='測定点')
    ax1.annotate(f'sigma_B {KV["A7075"]["sB"]:.1f} MPa',
                 xy=(A7075_ENG[7][0]*100,KV['A7075']['sB']),
                 xytext=(A7075_ENG[7][0]*100-2.5,KV['A7075']['sB']*1.05),
                 fontsize=8,color='darkred',arrowprops=dict(arrowstyle='->',color='darkred',lw=1))
    ax1.axhline(KV['A7075']['sy_lit'],color='gray',ls=':',lw=1)
    ax1.text(0.5,KV['A7075']['sy_lit']+20,f'0.2%耐力（文献）503 MPa',fontsize=8,color='gray')
    ax1.axvline(KV['A7075']['delta'],color='red',ls='--',lw=1,
                label=f'破断 delta={KV["A7075"]["delta"]:.2f}%')
    ax1.set_xlabel('公称ひずみ e [%]',fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]',fontsize=11)
    ax1.set_title('A7075（Al合金）\n公称応力-公称ひずみ線図（実測）',fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True,ls='--',alpha=0.5)
    et,st=zip(*A7075_TRUE); et,st=np.array(et),np.array(st)
    etx,stx=_smooth(et,st)
    ax2.plot(etx*100,stx,'darkorange',lw=2)
    ax2.plot(et*100,st,'o',color='darkorange',ms=5,label='測定点')
    ax2.set_xlabel('対数ひずみ epsilon [%]',fontsize=11)
    ax2.set_ylabel('真応力 σt [MPa]',fontsize=11)
    ax2.set_title('A7075（Al合金）\n真応力-対数ひずみ線図（実測）',fontsize=11)
    ax2.legend(fontsize=9); ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True,ls='--',alpha=0.5)
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def plot_fc250():
    fig,(ax1,ax2)=plt.subplots(1,2,figsize=(13,5.5))
    e,s=zip(*FC250_ENG); e,s=np.array(e),np.array(s)
    ax1.plot(e*100,s,'steelblue',lw=2)
    ax1.plot(e*100,s,'o',color='steelblue',ms=5,label='測定点（参考）')
    ax1.annotate(f'sigma_B {KV["FC250"]["sB"]:.1f} MPa',
                 xy=(FC250_ENG[-1][0]*100,KV['FC250']['sB']),
                 xytext=(FC250_ENG[-1][0]*100-0.3,KV['FC250']['sB']*1.08),
                 fontsize=8,color='darkred',arrowprops=dict(arrowstyle='->',color='darkred',lw=1))
    ax1.set_xlabel('公称ひずみ e [%]',fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]',fontsize=11)
    ax1.set_title('FC250（ねずみ鋳鉄）\n公称応力-公称ひずみ線図（参考データ）',fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True,ls='--',alpha=0.5)
    ax1.text(0.02,0.03,'注：CSV上 epsilon_max=1.08%，実測 delta=4.84%（寸法計算値）',
             transform=ax1.transAxes,fontsize=7,color='gray')
    et,st=zip(*FC250_TRUE); et,st=np.array(et),np.array(st)
    ax2.plot(et*100,st,'darkorange',lw=2)
    ax2.plot(et*100,st,'o',color='darkorange',ms=5,label='測定点（参考）')
    ax2.set_xlabel('対数ひずみ epsilon [%]',fontsize=11)
    ax2.set_ylabel('真応力 σt [MPa]',fontsize=11)
    ax2.set_title('FC250（ねずみ鋳鉄）\n真応力-対数ひずみ線図（参考データ）',fontsize=11)
    ax2.legend(fontsize=9); ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True,ls='--',alpha=0.5)
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def _hollomon_ss400():
    """Hollomon fit for SS400 plastic region (after Luders band to peak)."""
    ss = SS400_ENG[19:38]
    e = np.array([p[0] for p in ss])
    s = np.array([p[1] for p in ss])
    et = np.log(1+e); st = s*(1+e)
    ep = et - st/206e3
    mask = ep > 0.001
    ep, st = ep[mask], st[mask]
    c = np.polyfit(np.log(ep), np.log(st), 1)
    n, K = c[0], np.exp(c[1])
    ln_fit = np.polyval(c, np.log(ep))
    r2 = 1 - np.sum((np.log(st)-ln_fit)**2)/np.sum((np.log(st)-np.mean(np.log(st)))**2)
    return ep, st, K, n, r2

def plot_hollomon():
    ep, st, K, n, r2 = _hollomon_ss400()
    fig, ax = plt.subplots(figsize=(7,5))
    ax.loglog(ep*100, st, 'o', color='steelblue', ms=6, label='SS400 実測点')
    ep_line = np.logspace(np.log10(ep[0]), np.log10(ep[-1]), 100)
    st_line = K * ep_line**n
    ax.loglog(ep_line*100, st_line, color='darkorange', lw=2,
              label=f'Hollomon則: $\\sigma_t = {K:.0f}\\,\\varepsilon_p^{{{n:.3f}}}$\n$R^2 = {r2:.4f}$')
    ax.set_xlabel('相当塑性ひずみ $\\varepsilon_p$ [%] (対数軸)', fontsize=11)
    ax.set_ylabel('真応力 $\\sigma_t$ [MPa] (対数軸)', fontsize=11)
    ax.set_title(f'SS400 Hollomon則 両対数プロット（実測データ）\nK = {K:.0f} MPa, n = {n:.3f}', fontsize=11)
    ax.legend(fontsize=10); ax.grid(True, which='both', ls='--', alpha=0.5)
    fig.tight_layout(); return fig_to_buf(fig)

def plot_ss400_comparison():
    fig,ax=plt.subplots(figsize=(10,6))
    e,s=zip(*SS400_ENG)
    ax.plot(np.array(e)*100,s,color='steelblue',lw=2,label='SS400 通常試験（実測）')
    # 追加実験は伸び計データ未取得のため概略で表示
    d=MAT['SS400_add']; E=d['E_lit']*1000; sy=d['sy_lit']; sB=d['sB_lit']
    ef=elong('SS400_add')/100
    e1=np.linspace(0,sy/E,50); s1=E*e1
    ep=np.linspace(1e-3,ef-sy/E,100); e2=sy/E+ep; s2=np.clip(540*ep**0.20,sy-10,sB)
    e3=np.array([e2[-1],ef]); s3=np.array([s2[-1],320])
    e_add=np.r_[e1,e2,e3]; s_add=np.r_[s1,s2,s3]
    ax.plot(e_add*100,s_add,color='darkorange',ls='--',lw=2,
            label='SS400 追加実験（伸び計，概略）')
    ax.set_xlabel('公称ひずみ e [%]',fontsize=12)
    ax.set_ylabel('公称応力 sigma [MPa]',fontsize=12)
    ax.set_title('SS400 通常試験 vs 追加実験',fontsize=12)
    ax.legend(fontsize=11); ax.set_xlim(0); ax.set_ylim(0); ax.grid(True,ls='--',alpha=0.5)
    ax.text(0.02,0.03,'注：追加実験は伸び計データ確定後に実測値で更新予定',
            transform=ax.transAxes,fontsize=9,color='gray')
    fig.tight_layout(); return fig_to_buf(fig)

# ============================================================
# 文書構築
# ============================================================
def build():
    doc=Document(); set_margins(doc)
    sty=doc.styles['Normal']
    sty.font.name='Times New Roman'; sty.font.size=Pt(10.5)
    sty.paragraph_format.line_spacing=Pt(18)

    _eq=[0]; _fn=[0]; _tn=[0]
    def eq(): _eq[0]+=1; return _eq[0]
    def fn(): _fn[0]+=1; return _fn[0]
    def tn(): _tn[0]+=1; return _tn[0]
    def B(text, ind=True):
        return para(doc, text, indent_mm=4 if ind else 0)

    # ===== 表紙 =====
    for _ in range(3): doc.add_paragraph()
    para(doc,'機械航空宇宙工学実験　報告書',
         align=WD_ALIGN_PARAGRAPH.CENTER,size=16,bold=True,sb=0,sa=8,indent_mm=0)
    para(doc,'引張試験・シャルピー衝撃試験・破面観察',
         align=WD_ALIGN_PARAGRAPH.CENTER,size=13,bold=True,sb=0,sa=30,indent_mm=0)
    for label,val in [('科目','機械航空宇宙工学実験'),('学科・学年','機械航空宇宙工学科　3年　6B'),
                      ('学籍番号',''),('氏名',''),('提出日','')]:
        para(doc,f'{label}：{val}',align=WD_ALIGN_PARAGRAPH.CENTER,size=12,sb=0,sa=5,indent_mm=0)
    doc.add_page_break()

    # ===== 第1章 =====
    chapter_title(doc,'第1章　破断実験（引張試験）')
    h1(doc,'1.1　目的')
    B('本実験は，コンピュータ制御・油圧サーボ式万能試験機（島津製作所 UH-300kNA）を用いて，'
      '一般構造用圧延鋼材（SS400），ねずみ鋳鉄（FC250），およびアルミニウム合金（A7075）の'
      '3 種類の試験片に対して静的引張荷重を負荷し，'
      '降伏応力，引張強さ，破断伸び，絞り，ヤング率などの機械的性質を測定・評価することを目的とする．'
      'また，各材料の延性・脆性の違いを定量的に比較し，破壊挙動との対応を考察する．')

    h1(doc,'1.2　原理または理論')
    h2(doc,'1.2.1　公称応力・公称ひずみ')
    B('引張試験において，荷重を P [N]，初期断面積を A0 [mm2]，'
      '初期標点間距離を l0 [mm] とすると，公称応力 sigma [MPa] は式(1)，'
      '公称ひずみ e [-] は式(2) でそれぞれ定義される．')
    eq1=eq(); equation(doc,'σ = P / A₀',eq1)
    eq2=eq(); equation(doc,'e = (l − l₀) / l₀',eq2)
    B('ここで l [mm] は変形後の標点間距離である．'
      '公称量は変形中の断面積変化を考慮しないため，大変形域では実際の応力・ひずみと乖離する．')

    h2(doc,'1.2.2　真応力および対数ひずみ')
    B('変形中の瞬間断面積 A を用いた真応力 σₜ [MPa] は式(3) で定義される．'
      '体積一定（Al = A₀l₀）が成立する均一変形域では，式(4)(5) により公称量から変換できる．')
    eq3=eq(); equation(doc,'σₜ = P / A',eq3)
    eq4=eq(); equation(doc,'σₜ = σ (1 + e)',eq4)
    eq5=eq(); equation(doc,'ε = ln(l / l₀) = ln(1 + e)',eq5)
    B('対数ひずみ ε は微小変形増分を積み重ねたスカラー量であり，参照形状に依存しないという利点がある 1)．'
      '公称ひずみ e が十分小さい弾性域では ε ≈ e が成立する．')

    h2(doc,'1.2.3　相当応力および相当塑性ひずみ')
    B('三次元応力状態における降伏判定には，ミーゼスの相当応力 σ̄ [MPa] を用いる（式(6)）．'
      '相当塑性ひずみ増分 dε̄ᵖ は式(7) で定義される 2)．'
      '一軸引張では σ₂=σ₃=0 であり，相当応力は真応力に，相当塑性ひずみは軸方向真塑性ひずみに一致する．')
    eq6=eq(); equation(doc,'σ̄ = √{ [(σ₁−σ₂)² + (σ₂−σ₃)² + (σ₃−σ₁)²] / 2 }',eq6)
    eq7=eq(); equation(doc,'dε̄ᵖ = √(2/3) · dεᵖᵢⱼ dεᵖᵢⱼ',eq7)

    h2(doc,'1.2.4　指数硬化則（べき乗硬化則）')
    B('塑性域の真応力 σₜ と相当塑性ひずみ εₚ の関係をべき乗で近似した指数硬化則（Hollomon 則）を式(8) に示す 2)．'
      '両対数変換すると式(9) の直線関係となり，K [MPa]（強度係数）と n [-]（加工硬化指数）を回帰評価できる．'
      'n が大きいほど変形が均一に分布しやすい（Considere 条件：均一伸び eu = n）．')
    eq8=eq(); equation(doc,'σₜ = K · εₚⁿ',eq8)
    eq9=eq(); equation(doc,'ln σₜ = ln K + n · ln εₚ',eq9)

    h2(doc,'1.2.5　実験書 8. の諸量の定義')
    B('引張試験から得られる機械的性質の代表的な諸量を式(10)〜式(16) に定義する．')
    eq10=eq(); equation(doc,'σy  = Py   / A₀          （上降伏点応力 [MPa]）',eq10)
    eq11=eq(); equation(doc,'σL  = PLv  / A₀          （下降伏点応力 [MPa]）',eq11)
    eq12=eq(); equation(doc,'σB  = Pmax / A₀          （引張強さ [MPa]）',eq12)
    eq13=eq(); equation(doc,'σT  = Pt   / Af          （破断応力 [MPa]）',eq13)
    eq14=eq(); equation(doc,'δ   = (lf − l₀) / l₀ × 100  [%]  （破断伸び）',eq14)
    eq15=eq(); equation(doc,'φ   = (A₀ − Af) / A₀ × 100  [%]  （絞り）',eq15)
    eq16=eq(); equation(doc,'E   = P · λ / (A₀ · l₀)  （ヤング率 [GPa]）',eq16)
    B('ここで Py は上降伏点荷重 [kN]，PLv は下降伏点荷重 [kN]，Pmax は最大荷重 [kN]，'
      'Pt は破断時荷重 [kN]，Af は破断後最小断面積 [mm2]，lf は破断後標点間距離 [mm]，'
      'λ は荷重 P に対応する変位 [mm] である．'
      'φ の算出には破断部最小径の実測値が必要であり，本実験では体積一定仮定による近似値（※）を用いた．')

    h2(doc,'1.2.6　材料の破壊様式')
    B('金属材料の破壊様式は延性破壊と脆性破壊に大別される 1)．'
      '延性破壊（SS400，A7075 等）では，破断前に顕著な塑性変形が生じ，'
      '第 2 相粒子を核としたボイドの発生・成長・合体を経て破断に至る（ボイドコアレッセンス機構）1)．'
      '脆性破壊（FC250 等）では，塑性変形をほとんど伴わずにき裂が急速に伝播し，'
      '平坦な破面（へき開面，リバーパターン）が形成される 1)．'
      'FC250 ではフレーク状黒鉛が応力集中源となり低延性を示す 4)．')

    h1(doc,'1.3　装置の説明')
    fn1=fn()
    B(f'島津万能試験機 UH-300kNA（Fig. {fn1}）を使用した．'
      'コンピュータ制御・油圧サーボ式で最大試験力 300 kN，試験力精度 ±1%（JIS B 7721 準拠）である．'
      '試験力はロードセルで計測し，荷重-変位曲線はチャートレコーダに自動記録される．')
    fig_caption(doc,fn1,'島津万能試験機 UH-300kNA（実験書 Fig. 1 参照）')

    h1(doc,'1.4　実験方法')
    h2(doc,'1.4.1　試験片の準備')
    B('マイクロメーターを用いて各試験片平行部の中心近傍において直径を 3 箇所測定し，'
      '平均直径 d から初期断面積 A₀ = π(d/2)² を算出した．'
      '標点間距離 l₀ はノギスで測定した．'
      '試験片は JIS Z 2241 の 4 号試験片形状であり，材質は SS400，FC250，A7075 の 3 種類である．')
    h2(doc,'1.4.2　引張試験手順')
    B('電源投入後，負荷ポンプ起動，ラム初期位置設定，荷重ゼロ・スパン調整を行った．'
      '試験片を上つかみに固定後，LOAD RANGE を 150 kN，応力速度 3.0 MPa/s に設定した．'
      '断面積入力・ZERO 確認後に下つかみに固定し，ENTER キーにより試験を開始した．'
      '破断後，ラムを初期位置に戻してから試験片を取り外し，破断部最小径測定と破面スケッチを行った．')

    h1(doc,'1.5　測定結果')
    B('Table 1 に各試験片の初期寸法と破断後標点間長さを示す．A0 は 3 回の直径測定値の平均から算出した．')
    tn1=tn()
    tbl_caption(doc,tn1,'引張試験試験片寸法および破断後標点間長さ')
    tbl=doc.add_table(rows=1,cols=7); tbl.style='Table Grid'; tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','d1 [mm]','d2 [mm]','d3 [mm]','l0 [mm]','A0 [mm2]','lf [mm]']):
        c=tbl.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    for rd in [('SS400（通常）',14.060,14.053,14.068,82.18,155.168,105.67),
               ('FC250',        14.060,13.992,14.022,62.22,154.491, 65.23),
               ('A7075',        13.982,14.017,14.023,62.45,154.099, 67.25)]:
        row=tbl.add_row().cells
        for i,v in enumerate(rd):
            row[i].text=str(v) if isinstance(v,str) else f'{v:.3f}'
            row[i].paragraphs[0].runs[0].font.size=Pt(9)
    B('')
    B('式(14) より算出した破断伸び delta を Table 2 に示す．'
      'SS400（28.58%）>> A7075（7.69%）> FC250（4.84%）の延性序列が明確に確認できる．')
    tn2=tn()
    tbl_caption(doc,tn2,'破断伸び delta（実測値）')
    tbl2=doc.add_table(rows=1,cols=4); tbl2.style='Table Grid'; tbl2.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','l0 [mm]','lf [mm]','delta [%]']):
        c=tbl2.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    for key,name in [('SS400','SS400'),('FC250','FC250'),('A7075','A7075')]:
        d=MAT[key]; row=tbl2.add_row().cells
        row[0].text=name; row[1].text=f'{d["l0"]:.2f}'
        row[2].text=f'{d["lf"]:.2f}'; row[3].text=f'{elong(key):.2f}'
        for c in row: c.paragraphs[0].runs[0].font.size=Pt(9)
    B('')

    h1(doc,'1.6　課題')
    h3(doc,'課題 1　実験書 8. の諸量について説明せよ．')
    B('式(10)〜式(16) で定義される機械的性質の諸量について，以下に物理的意味を述べる．')
    B('（1）降伏応力 σy（式(10)）：弾性変形から塑性変形へ移行する際の応力．'
      'SS400 では荷重が急激に低下するときの上降伏点荷重 Py を用いる．'
      '明確な降伏点を示さない材料（A7075 等）では永久伸び 0.2% に相当する耐力を代用する．')
    B('（2）下降伏点応力 σL（式(11)）：SS400 特有の現象であり，上降伏点通過後に荷重が低下して安定した応力値．'
      'この応力での水平域（リューダース帯の伝播域）が続いた後，加工硬化が始まる．')
    B('（3）引張強さ σB（式(12)）：最大荷重 Pmax を原断面積 A₀ で除した値．'
      '材料強度の代表指標として広く使用される．最大荷重到達後にくびれが始まる．')
    B('（4）破断応力 σT（式(13)）：破断時荷重 Pt を破断後最小断面積 Af で除した値．'
      '延性材料ではくびれにより Af < A₀ となるため σT > σB となることが多い．')
    B('（5）破断伸び δ（式(14)）：試験前後の標点間距離変化率で延性指標の一つ．'
      '標点間距離の設定によって値が変わるため JIS 規格試験片での測定が必要である．')
    B('（6）絞り φ（式(15)）：破断部断面積減少率で延性指標の一つ．'
      'くびれが生じない脆性材料では φ ≈ 0 となる．正確な値には破断部最小径の実測が必要であり，'
      '本実験では体積一定仮定の近似値を用いた（※体積一定仮定によるためネッキング後の誤差を含む）．')
    B('（7）ヤング率 E（式(16)）：弾性域における応力とひずみの比例係数．材料固有の剛性を表す．'
      '弾性域の荷重変化量 P，対応する変位 λ，および A₀，l₀ から算出する．')

    h3(doc,'課題 2　真応力・対数ひずみについて，式を交えて説明せよ．')
    B('公称応力（式(1)）は初期断面積 A₀ で除するため，変形が進むと実際の応力より小さくなる．'
      '真応力（式(3)）は瞬間断面積 A を用いるため，変形とともに増加する実際の応力状態を反映する．'
      '体積一定（Al = A₀l₀）が成立する均一変形域では式(4)(5) により公称量から変換できる．'
      'ネッキング開始後は体積一定仮定が成立しなくなり，破断部最小断面積の実測が必要となる．'
      '本実験の SS400（真応力-対数ひずみ線図；Fig. 4 右）では最大荷重点（δ = 26.72%）まで変換を適用した．')
    B('対数ひずみ（式(5)）は微小変形増分 dl/l を積分した量であり，参照形状に依存しないという特徴を持つ 1)．'
      '加工硬化のモデリング，多軸変形評価，有限要素解析における材料則の記述に広く用いられる．'
      '公称ひずみ e が小さい弾性域では ε ≈ e が成立し，両者は実用上一致する．')

    h3(doc,'課題 3　相当応力・相当塑性ひずみについて，式を交えて説明せよ．')
    B('三次元応力場における降伏を一つのスカラー量で表すために相当量が定義される．'
      'ミーゼスの相当応力（式(6)）は多軸応力場における降伏を等価な一軸引張応力として評価する量であり，'
      '式の値が降伏応力 σy に達したときに降伏が開始するとするのがミーゼス降伏条件である 2)．')
    B('相当塑性ひずみ増分（式(7)）は塑性ひずみ増分テンソルの不変量として定義され，'
      '変形量の大きさを表すスカラー量である．'
      '一軸引張では σ₁=σ，σ₂=σ₃=0 であるから，'
      '相当応力は真応力に，相当塑性ひずみは軸方向真塑性ひずみと等価になる 2)．'
      'これにより，一軸引張試験から得た真応力-対数ひずみ曲線（弾性部を除く）は'
      '相当応力-相当塑性ひずみ関係として直接利用できる．')

    h3(doc,'課題 4　指数硬化則（べき乗硬化則）について，式を交えて説明せよ．')
    _ep, _st, _K, _n, _r2 = _hollomon_ss400()
    B('加工硬化を記述するモデルとして式(8) の指数硬化則（Hollomon 則）が広く用いられる 2)．'
      '強度係数 K は εₚ=1 における真応力の外挿値，'
      '加工硬化指数 n は硬化の速さを示す無次元指数である．'
      '式(9) の両対数変換により塑性変形域のデータが直線上に並ぶため，'
      '引張試験の真応力-対数ひずみ線図から K と n を回帰評価できる．'
      'n が大きいほど変形が均一に分布しやすく，ネッキング開始における均一伸び eu=n の関係がある（Considere の条件）．')
    B(f'実測データ（SS400 リューダース帯以降〜最大荷重点，{len(_ep)} 点）に対して最小二乗回帰を行った結果，'
      f'K = {_K:.1f} MPa，n = {_n:.3f}，R2 = {_r2:.4f} が得られた．'
      f'Considere の均一伸び eu = n = {_n*100:.1f}% は実測の最大荷重点まで均一伸び {SS400_ENG[37][0]*100:.2f}% と概ね整合する．'
      f'A7075 については，クロスヘッド変位に機械コンプライアンス（見かけ E ≈ 14 GPa，実 E = 71.7 GPa）が含まれるため，'
      f'塑性ひずみの精度が低く Hollomon パラメータの確度ある算出には伸び計データが必要である．')

    fn_holl = fn()
    B(f'Fig. {fn_holl} に SS400 の両対数プロット（Hollomon 則フィット）を示す．')
    insert_figure(doc, plot_hollomon(), width_cm=11)
    fig_caption(doc, fn_holl, f'SS400 Hollomon 則両対数プロット（実測データ，K={_K:.0f} MPa，n={_n:.3f}）')
    B('')

    h3(doc,'課題 5　材料の破壊様式について説明せよ．')
    B('金属材料の破壊様式は延性破壊と脆性破壊に大別される 1)．')
    B('延性破壊（SS400，A7075 等）では，巨視的には顕著なくびれと大きな塑性変形を伴う．'
      'ミクロ的には，第 2 相粒子（介在物，析出物）を核としてボイドが発生・成長し，'
      '合体によって破断に至る（ボイドコアレッセンス機構）1)．'
      '破面にはディンプル（半球状の窪み）が多数観察される．'
      '引張軸に垂直な中央繊維状領域と外周の 45° せん断リップ（カップコーン破面）が特徴的である 1)．')
    B('脆性破壊（FC250 等）では塑性変形がきわめて小さい状態でき裂が急速に伝播する．'
      'FC250 ではフレーク状黒鉛の先端が応力集中源となり，引張荷重に対してほとんど変形なく破断する 4)．'
      '破面は平坦であり，へき開面やリバーパターンが形成されると考えられる 1)．'
      'A7075 は高強度 Al 合金であるが，一般に延性的破壊形態を示し，析出物を核とした微細ディンプルが'
      '形成されると考えられる 5)．')

    h3(doc,'課題 6　各材料の公称応力-公称ひずみ線図および真応力-対数ひずみ線図を作成せよ．')
    fn_all=fn()
    fn_ss=fn(); fn_a7=fn(); fn_fc=fn()
    B(f'Fig. {fn_all} に 3 材料比較の公称応力-公称ひずみ線図を示す（実測データ）．'
      f'Fig. {fn_ss} に SS400，Fig. {fn_a7} に A7075，Fig. {fn_fc} に FC250 の'
      '公称応力-公称ひずみ線図（左）および真応力-対数ひずみ線図（右）を示す．'
      '真応力-対数ひずみ線図は体積一定仮定が成立する均一変形域（最大荷重点まで）に限定して変換した．')
    insert_figure(doc, plot_overview())
    fig_caption(doc,fn_all,'3 材料比較 公称応力-公称ひずみ線図（実測データ）')
    B('')
    insert_figure(doc, plot_ss400(), width_cm=15)
    fig_caption(doc,fn_ss,'SS400（構造用鋼） 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（実測）')
    B('')
    insert_figure(doc, plot_a7075(), width_cm=15)
    fig_caption(doc,fn_a7,'A7075（Al合金） 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（実測）')
    B('')
    insert_figure(doc, plot_fc250(), width_cm=15)
    fig_caption(doc,fn_fc,'FC250（ねずみ鋳鉄） 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（参考データ）')
    B('')
    B(f'SS400 の公称線図（Fig. {fn_ss}）では，上降伏点（{KV["SS400"]["sy"]:.1f} MPa）における荷重低下，'
      f'下降伏点（{KV["SS400"]["sL"]:.1f} MPa）付近での水平域（リューダース伸び），加工硬化，'
      f'引張強さ（{KV["SS400"]["sB"]:.1f} MPa）到達後のくびれによる荷重低下，および破断が確認できる．'
      f'A7075 では明瞭な降伏点は存在せず，0.2% 耐力（文献値 ≈ {KV["A7075"]["sy_lit"]} MPa）付近から塑性変形が始まる．'
      f'FC250 では線形的な応力増加の後，引張強さ（{KV["FC250"]["sB"]:.1f} MPa）と同時にほぼ突然破断し，脆性的挙動が確認できる．')

    h3(doc,'課題 7　実験書 8. の諸量算出およびヤング率の比較考察')
    B('実測値から算出した諸量を以下に示す．破断伸び delta は Table 2 に示したとおりである．'
      '絞り phi は体積一定仮定に基づく近似値であり，正確な値は破断部最小径の実測を要する．')
    # 実測値テキスト
    B(f'SS400（実測値）：上降伏点応力 σy = {KV["SS400"]["sy"]:.1f} MPa（Py = {KV["SS400"]["sy"]*MAT["SS400"]["A0"]/1000:.1f} kN），'
      f'下降伏点応力 σL = {KV["SS400"]["sL"]:.1f} MPa，'
      f'引張強さ σB = {KV["SS400"]["sB"]:.1f} MPa（Pmax = {KV["SS400"]["Pmax"]:.1f} kN），'
      f'δ = {KV["SS400"]["delta"]:.2f}%，φ ≈ {KV["SS400"]["phi"]:.2f}%（体積一定仮定による近似）．')
    B(f'FC250（参考データ）：脆性材料のため明確な降伏点なし．'
      f'σB = {KV["FC250"]["sB"]:.1f} MPa（Pmax = {KV["FC250"]["Pmax"]:.1f} kN），'
      f'δ = {KV["FC250"]["delta"]:.2f}%，φ ≈ {KV["FC250"]["phi"]:.2f}%（体積一定仮定）．'
      f'JIS G 5501 規定の最小引張強さ 250 MPa を上回っている．')
    B(f'A7075（実測値）：σy は 0.2% 耐力で定義されるが，クロスヘッド変位には機械コンプライアンスが含まれるため'
      f'今回の測定からは算出不可（文献値 503 MPa）．'
      f'σB = {KV["A7075"]["sB"]:.0f} MPa（Pmax = {KV["A7075"]["Pmax"]:.1f} kN），'
      f'δ = {KV["A7075"]["delta"]:.2f}%，φ ≈ {KV["A7075"]["phi"]:.2f}%（体積一定仮定）．')
    tn3=tn()
    tbl_caption(doc,tn3,'引張試験 機械的性質まとめ（実測値・文献値）')
    tbl3=doc.add_table(rows=1,cols=6); tbl3.style='Table Grid'; tbl3.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','σy [MPa]','σB [MPa]','E 文献[GPa]','δ 実[%]','φ [%]※']):
        c=tbl3.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    tbl3_data = [
        ('SS400', f'{KV["SS400"]["sy"]:.1f}（上降伏点，実測）', f'{KV["SS400"]["sB"]:.1f}（実測）',
         str(KV["SS400"]["E_lit"]), f'{KV["SS400"]["delta"]:.2f}', f'{KV["SS400"]["phi"]:.2f}'),
        ('FC250', '− （脆性，降伏点なし）', f'{KV["FC250"]["sB"]:.1f}（参考データ）',
         str(KV["FC250"]["E_lit"]), f'{KV["FC250"]["delta"]:.2f}', f'{KV["FC250"]["phi"]:.2f}'),
        ('A7075', '503（文献値，0.2%耐力）', f'{KV["A7075"]["sB"]:.0f}（実測）',
         str(KV["A7075"]["E_lit"]), f'{KV["A7075"]["delta"]:.2f}', f'{KV["A7075"]["phi"]:.2f}'),
    ]
    for row_data in tbl3_data:
        row=tbl3.add_row().cells
        for i,v in enumerate(row_data):
            row[i].text=v
            row[i].paragraphs[0].runs[0].font.size=Pt(9)
    B('※ φ は体積一定仮定による近似値．正確な値には破断部最小径の実測が必要．',ind=False)
    B('')
    B('ヤング率の文献値は SS400 で 206 GPa，A7075 で 71.7 GPa，FC250 で 130 GPa 程度である 2) 4) 5)．'
      '今回のクロスヘッド変位には機械コンプライアンスが含まれるため，'
      '弾性域から直接算出した E は文献値より過小評価となる（2.6 節参照）．'
      '追加実験（LVDT 伸び計）のデータが得られれば，コンプライアンス補正により正確な E との比較が可能になる．')
    B('破断伸びの比較から SS400（28.58%）>> A7075（7.69%）> FC250（4.84%）の延性序列が確認される．'
      'SS400 の高延性は明瞭な降伏プラトーと広い加工硬化域に起因し，'
      'FC250 の低延性はフレーク黒鉛による応力集中に起因する 4)．'
      'A7075 の中程度延性は MgZn2 析出物による析出強化と高強度ゆえの小さな均一伸びに起因する 5)．')

    h1(doc,'1.7　考察')
    # Toughness calc
    import numpy as _np
    _U_SS = _np.trapz([p[1] for p in SS400_ENG], [p[0] for p in SS400_ENG])
    _U_A7 = _np.trapz([p[1] for p in A7075_ENG], [p[0] for p in A7075_ENG])
    _U_FC = _np.trapz([p[1] for p in FC250_ENG], [p[0] for p in FC250_ENG])
    B(f'破断伸び delta の序列は SS400（{KV["SS400"]["delta"]:.2f}%）>> '
      f'A7075（{KV["A7075"]["delta"]:.2f}%）> FC250（{KV["FC250"]["delta"]:.2f}%）であり，'
      'SS400 と FC250 の間には約 6 倍の差が存在する．')
    B(f'公称応力-公称ひずみ線図の面積（公称靭性 Ur）を台形積分で算出すると，'
      f'SS400: {_U_SS:.1f} MJ/m3，A7075: {_U_A7:.1f} MJ/m3，FC250: {_U_FC:.1f} MJ/m3 となった．'
      f'SS400 は FC250 の約 {_U_SS/_U_FC:.0f} 倍のエネルギー吸収能を示し，延性差が靭性に直結することが定量的に確認できる．'
      '強度は A7075 が最大でありながら，低延性（7.69%）に起因して靭性は SS400 を大きく下回る．')
    B(f'SS400 の Hollomon 則パラメータ（K = {_K:.0f} MPa，n = {_n:.3f}）は'
      '真応力-塑性ひずみ関係に良く適合する（R2 = {_r2:.4f}）．'
      f'n = {_n:.3f} は炭素鋼典型値（0.20〜0.27 程度）よりやや高く，'
      'これは本試験片の降伏後に広いリューダース伸びが存在し，その後の加工硬化域が長いことに起因する可能性がある．')
    B('真応力-対数ひずみ線図では，均一変形域全体にわたって真応力は公称応力より高い値を示す．'
      'ネッキング開始（最大荷重点）以降は体積一定仮定が成立しないため，'
      '破断部最小径の実測値を用いた補正が必要となる．'
      'ヤング率については，伸び計を用いた追加実験データが得られれば，'
      'コンプライアンス誤差を排除した正確な値が算出できる 6)．')

    # ===== 第2章 =====
    doc.add_page_break()
    chapter_title(doc,'第2章　追加実験（伸び計を用いた実験）')
    h1(doc,'2.1　目的')
    B('本実験は，LVDT（直動差動変圧器）伸び計を試験片平行部に装着し，'
      '試験機クロスヘッド変位に含まれる機械コンプライアンス誤差を排除して，'
      'SS400 のヤング率を精密に計測することを目的とする．'
      'また，伸び計の有無による応力-ひずみ線図の相違を考察し，コンプライアンス補正の根拠を明確にする．')

    h1(doc,'2.2　原理または理論')
    h2(doc,'2.2.1　LVDT 伸び計の動作原理')
    B('作動変圧式伸び計は LVDT（Linear Variable Differential Transformer，直動差動変圧器）を'
      '測定素子として利用する 6)．'
      'LVDT は一次コイルと二つの二次コイルを同軸に配置した構造であり，'
      '中心の可動鉄心の位置に応じて二次コイル間の誘起電圧差が変化する．'
      '一次コイルに交流電圧を印加して磁束を発生させ，可動鉄心の変位量に比例した差動電圧を'
      '検出・増幅することで試験片の伸びを高精度に計測する 6)．'
      '非接触型電気変換方式のため，機械的摩擦の影響がなく応答性・繰り返し精度に優れる．')
    h2(doc,'2.2.2　機械コンプライアンスの概念')
    B('試験機クロスヘッド変位 lambda_cross には，試験片平行部の伸び lambda_spec に加えて，'
      '試験機フレームの弾性変形 lambda_frame，つかみ部の局所変形 lambda_grip 等が加算される．'
      '弾性域での計測精度が特に問題となり，LVDT 伸び計を用いることで lambda_spec のみを抽出できる．'
      '式(16) を用いたヤング率算出において，コンプライアンス補正の有無による差が定量的に評価できる．')

    h1(doc,'2.3　装置の説明')
    B('LVDT 伸び計を使用した．仕様：標点間距離 50.00 mm，計測最大伸び 10%（5.00 mm），'
      '目盛り 1 目量 0.5 mm（1 V 相当）．計測機器の 2 号端子に接続した．'
      '引張試験機本体は第1章 1.3 節に示した島津万能試験機 UH-300kNA を使用した．')

    h1(doc,'2.4　実験方法')
    B('SS400（平行部長さ 80 mm，l0 = 82.30 mm，A0 = 154.165 mm2）を用いて引張試験を行った．'
      'LVDT 伸び計を試験片平行部に装着し，計測範囲（5.00 mm）到達前に取り外した後，'
      '試験を継続して最大荷重点まで記録した（伸び計保護のため最大荷重点到達後は中断）．'
      '試験力設定等は第1章 1.4.2 節の手順と同様である．')

    h1(doc,'2.5　測定結果')
    B('Table 4 に SS400 追加実験試験片の寸法を示す．'
      '通常試験の試験片（Table 1）と同一材料であるが，平行部長さが異なる（80 mm，JIS 2241 14A 号相当）．')
    tn4=tn()
    tbl_caption(doc,tn4,'SS400 追加実験試験片寸法')
    tbl4=doc.add_table(rows=1,cols=6); tbl4.style='Table Grid'; tbl4.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','d1 [mm]','d2 [mm]','d3 [mm]','l0 [mm]','A0 [mm2]']):
        c=tbl4.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    d=MAT['SS400_add']
    row=tbl4.add_row().cells
    for i,v in enumerate(['SS400（追加）',d['d1'],d['d2'],d['d3'],d['l0'],d['A0']]):
        row[i].text=str(v) if isinstance(v,str) else f'{v:.3f}'
        row[i].paragraphs[0].runs[0].font.size=Pt(9)
    B('')
    fn_add=fn()
    B(f'Fig. {fn_add} に SS400 通常試験と追加実験の公称応力-公称ひずみ線図の比較を示す（追加実験は概略）．'
      '追加実験では最大荷重点まで伸び計データが取得され，弾性域の傾きを精密に評価できる．')
    insert_figure(doc, plot_ss400_comparison())
    fig_caption(doc,fn_add,'SS400 通常試験（実測）と追加実験（概略）公称応力-公称ひずみ線図の比較')
    B('')

    h1(doc,'2.6　課題')
    h3(doc,'課題 8　伸び計（作動変圧式）の測定原理を調べ，説明せよ．')
    B('2.2.1 節に示したように，LVDT を用いた伸び計は一次コイルへの交流印加で磁束を発生させ，'
      '可動鉄心の変位に比例した差動電圧を検出することで試験片の伸びを高精度に計測する 6)．'
      '本実験使用の伸び計の仕様は，標点間距離 50.00 mm，計測最大伸び 10%（5.00 mm），'
      '目盛り 1 目量 0.5 mm（1 V 相当）である．')

    h3(doc,'課題 9　追加実験の結果からヤング率を計算し，破断実験の結果や一般的な値と比較し，考察せよ．')
    B('追加実験（SS400，l₀ = 82.30 mm，A₀ = 154.165 mm²）では，'
      '弾性域における荷重 P と伸び計出力変位 λ の比例関係から式(16) によりヤング率を算出する．'
      'E = P · λ / (A₀ · l₀) において弾性域の傾き ΔP/Δλ を読み取り E を求める．'
      '（実験当日の計測グラフから直接値を記入する．）')
    B('文献値は E = 206 GPa であり，伸び計を用いた試験では機械コンプライアンスが排除されるため'
      '文献値に近い値が得られると期待される．'
      '一方，クロスヘッド変位を用いた通常試験では，フレームおよびつかみ部の弾性変形が加算されるため'
      'ヤング率が過小評価される傾向がある 6)．両者の差からコンプライアンスを定量化できる．')

    h3(doc,'課題 10　最大荷重時の公称応力と真応力を求めよ．')
    d_add=MAT['SS400_add']; Pmax_add=KV['SS400']['sB_lit']*d_add['A0']/1000
    B(f'追加実験（SS400，A₀ = {d_add["A0"]:.3f} mm²）では，'
      '最大荷重時の公称応力 σB = Pmax / A₀，真応力 σₜ(max) = σB × (1 + emax) を算出する．'
      '（emax は最大荷重時公称ひずみ = λmax / l₀）'
      f'文献値ベースの概算では Pmax ≈ {Pmax_add:.1f} kN，σB ≈ 400 MPa，'
      '均一伸び eu ≈ n ≈ 0.20 より σₜ(max) ≈ 400 × 1.20 = 480 MPa 程度と見積もれる．')

    h3(doc,'課題 11　標点間体積一定の仮定のもとで真応力-対数ひずみ関係を表わせ．')
    B(f'追加実験の荷重-変位データ（l₀ = {d_add["l0"]:.2f} mm，A₀ = {d_add["A0"]:.3f} mm²）から，'
      '各測定点において公称量 e = λ/l₀，σ = P/A₀ を算出し，'
      '体積一定（Al = A₀l₀）の仮定を用いて式(4)(5) により真量へ変換する．'
      '変換が有効な均一変形域（最大荷重点まで）のデータをプロットして真応力-対数ひずみ線図を作成する．')

    h3(doc,'課題 12　伸び計使用の有無による応力-ひずみ関係の違いについて考察せよ．')
    B('伸び計使用時（平行部標点間変位を直接計測）と未使用時（クロスヘッド変位使用）では，'
      '次の 2 要因によって応力-ひずみ関係に違いが生じる．')
    B('(1) 機械コンプライアンスの影響：クロスヘッド変位にはフレームおよびつかみ部の弾性変形が含まれるため，'
      '弾性域においてヤング率が過小評価される．塑性域ではこの寄与が相対的に小さくなり差が縮まる傾向にある 6)．')
    B('(2) つかみ部局所変形の影響：テーパー移行部での変形が全体の伸びに混入する．'
      '伸び計は平行部を直接計測するためこの影響を排除でき，材料本来の弾性係数をより正確に反映する．')

    h1(doc,'2.7　考察')
    B('伸び計を用いることで機械コンプライアンスの影響を排除し，'
      'SS400 のヤング率を文献値（206 GPa）に近い値で評価できると期待される．'
      'クロスヘッド変位を用いた場合の過小評価量からコンプライアンスを定量化することで，'
      '今後の通常試験データに対するヤング率補正の根拠が得られる．')
    B(f'SS400 と追加実験試験片の公称応力-公称ひずみ線図を比較すると（Fig. {fn_add}），'
      '試験片寸法のわずかな差（平行部長さ 80 mm vs 82.18 mm，A0 の微差）による影響は小さく，'
      '破断伸びの差は試験片内の標点位置の違いに起因すると考えられる．')

    # ===== 第3章 =====
    doc.add_page_break()
    chapter_title(doc,'第3章　シャルピー衝撃試験')
    h1(doc,'3.1　目的')
    B('本実験は，振り子型シャルピー試験機を用いて SS400，FC250，A7075 の 3 材料に対して'
      '衝撃荷重を負荷し，シャルピー吸収エネルギー K および衝撃値 rho を算出することを目的とする．'
      '各材料の靭性を定量的に評価し，静的引張試験の結果との比較・考察を行う．')

    h1(doc,'3.2　原理または理論')
    h2(doc,'3.2.1　シャルピー衝撃試験の原理')
    B('振り子型ハンマを最大持上げ角 alpha から放ち，試験片打撃後の振上がり角 beta から吸収エネルギーを算出する．'
      'シャルピー吸収エネルギー K [J] は式(17)，シャルピー衝撃値 rho [J/cm2] は式(18) で求める 3)．')
    eq17=eq(); equation(doc,'KV = W · r (cos β − cos α) − L',eq17)
    eq18=eq(); equation(doc,'ρ = KV / A',eq18)
    B(f'ここで W = {W_kg} kg，g = 9.807 m/s²，r = {r_m} m，α = {alpha}°（最大持上げ角），'
      f'L [J] はエネルギー損失（試験片なし 100 往復から算出），A [cm²] はノッチ部断面積である 3)．'
      f'cos α = cos {alpha}° = {cos_a:.4f}．W · r = {W_N:.2f} × {r_m} = {Wr:.2f} J．')

    h1(doc,'3.3　装置の説明')
    fn_charpy=fn()
    B(f'振り子型シャルピー試験機（Fig. {fn_charpy}）を使用した．'
      f'ハンマ質量 W = {W_kg} kg，重心半径 r = {r_m} m，最大持上げ角 alpha_m = {alpha} deg である．')
    fig_caption(doc,fn_charpy,'シャルピー衝撃試験機 概観図（実験書参照）')

    h1(doc,'3.4　実験方法')
    B('マイクロメーターおよびノギスで試験片の l，b，h1，h2 を測定し，'
      'ノッチ部断面積 A = b * h2 を算出した後，試験片を位置決めして取り付けた．'
      'ハンマを最大持上げ角 alpha_m ≈ 146.5 deg まで持ち上げて解放し，打撃後の振上がり角 beta を読み取った．'
      'また，試験片なしで 100 往復させてエネルギー損失 L を算出した．')

    h1(doc,'3.5　測定結果')
    B('Table 5 にシャルピー試験片の寸法を示す．ノッチ部断面積 A = b * h2 で算出した．')
    tn5=tn()
    tbl_caption(doc,tn5,'シャルピー試験片寸法')
    tbl5=doc.add_table(rows=1,cols=7); tbl5.style='Table Grid'; tbl5.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','l [mm]','b [mm]','h1 [mm]','h2 [mm]','A [mm2]','A [cm2]']):
        c=tbl5.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    for mat,d in CHARPY.items():
        row=tbl5.add_row().cells
        for i,v in enumerate([mat,d['l'],d['b'],d['h1'],d['h2'],d['A'],d['A']/100]):
            row[i].text=str(v) if isinstance(v,str) else (f'{v:.4f}' if v<1 else f'{v:.2f}')
            row[i].paragraphs[0].runs[0].font.size=Pt(9)
    B('')

    h1(doc,'3.6　課題')
    h3(doc,'課題 13　材料の特徴と吸収エネルギーの関係について考察せよ．')
    B(f'式(17) に定数を代入すると，KV = {Wr:.2f}(cos β + {-cos_a:.4f}) − L [J] となる．'
      'β の読取り値と摩擦損失 L を代入すれば KV が算出できる．'
      'Table 6 に試験片寸法から算出したノッチ部断面積 A と，β 測定値記入欄を示す．')
    tn6=tn()
    tbl_caption(doc,tn6,'シャルピー試験結果（β 測定値・KV・ρ は実験当日の読取り値を記入）')
    tbl6=doc.add_table(rows=1,cols=5); tbl6.style='Table Grid'; tbl6.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['材料','β [deg]','KV [J]','A [cm²]','ρ [J/cm²]']):
        c=tbl6.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.size=Pt(9)
    for mat,d in CHARPY.items():
        row=tbl6.add_row().cells
        row[0].text=mat; row[1].text=''; row[2].text=''
        row[3].text=f'{d["A"]/100:.3f}'; row[4].text=''
        for c in row: c.paragraphs[0].runs[0].font.size=Pt(9)
    B('※ β，KV，ρ の欄は読取り値取得後に手書きで記入すること．',ind=False)
    B('')
    B('材料の靭性は応力-ひずみ曲線下の面積（単位体積あたりの破断エネルギー）に対応し，強度と延性の積として評価できる．'
      f'引張試験から算出した公称靭性（台形積分）は SS400 が約 114 MJ/m³ であり，'
      f'A7075（{50:.0f} MJ/m³）・FC250（1.6 MJ/m³）を大きく上回る．'
      'よって SS400 が最大，FC250 が最小の吸収エネルギーを示すと予測され，'
      'A7075 は中間と予想される．β 角測定値が得られた後，KV・ρ を算出して定量的に確認する．')

    h3(doc,'課題 14　脆性破面率と吸収エネルギーの関係について考察せよ．')
    B('脆性破面率 B [%] は，破面全面積に対する脆性破面面積の比率（B = C/A * 100）であり，'
      '材料の破壊様式を定量化する指標の一つである．')
    B('一般に B が大きい材料ほどシャルピー吸収エネルギーが小さい傾向がある 1)．'
      'これは延性的な破壊過程（ボイドの発生・成長）ではき裂先端に塑性変形帯が発達し，'
      'エネルギー吸収量が増加するためである．'
      'FC250 は黒鉛フレークを起点とした脆性的破壊が支配的であり B が高く吸収エネルギーが小さい 1)．'
      'SS400 では破面の大部分が延性的（ディンプル）であるため B が低く吸収エネルギーが大きい 1)．'
      'β 角測定後に KV・ρ を算出することで，B と ρ の定量的相関を考察できる．')

    h1(doc,'3.7　考察')
    B(f'静的引張試験で得た延性序列（SS400 >> A7075 > FC250）は，'
      'シャルピー試験での吸収エネルギーの序列とも概ね対応すると考えられる．'
      'これは強度と延性の積が靭性の尺度であるという観点と整合する．'
      'beta 角の精度が K の計算値に直接影響するため，角度読取り精度の確認が重要である．')
    B('また，脆性破面率 B とシャルピー衝撃値 rho の対応は，引張試験で確認した破壊様式（第1章 1.2.6）'
      'と定量的に整合することが予想される．'
      'FC250 は引張試験でほぼ脆性破壊を示したことからも B が高く rho が小さいという予測は妥当である 4)．')


    # ===== 第4章 =====
    doc.add_page_break()
    chapter_title(doc,'第4章　破面観察')

    SEM_DIR = "/Users/takaoka/report/sem_images"

    def insert_sem_pair(doc, path1, path2):
        """SEM画像2枚を横並びで挿入する（2列テーブル使用）"""
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl._tbl.insert(0, tblPr)
        tblBorders = OxmlElement('w:tblBorders')
        for border in ['top','left','bottom','right','insideH','insideV']:
            b = OxmlElement(f'w:{border}')
            b.set(qn('w:val'), 'none')
            tblBorders.append(b)
        tblPr.append(tblBorders)
        for i, path in enumerate([path1, path2]):
            cell = tbl.rows[0].cells[i]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(path, width=Cm(7.0))
        return tbl

    import os as _os

    def sem(mat, region, mag):
        return _os.path.join(SEM_DIR, f"{mat}{region}{mag}.jpg")

    REGION_JP = {
        'c':  '中央繊維状域',
        'CC': 'カップ中央部',
        'CN': 'カップ-コーン遷移域',
        'CO': 'カップ外縁部',
        'E':  'せん断リップ',
    }
    REGION_JP_FC = {
        'c':  '中央繊維状域',
        'CC': 'カップ中央部',
        'CN': '中央脆性域',
        'CO': '外縁脆性域',
        'E':  'へき開域',
    }

    h1(doc,'4.1　目的')
    B('本実験は，引張試験後の破断面を走査型電子顕微鏡（SEM）で観察し，'
      '各材料（SS400，FC250，A7075）の延性・脆性・混合破壊に特有のミクロ形態を確認することを目的とする．'
      'SS400（延性破壊），FC250（脆性破壊），A7075（混合破壊）の 3 材料について'
      '破面の各部位（中央繊維状域，カップ中央部，遷移域，外縁部，せん断リップ）を'
      'それぞれ ×100・×500 の 2 倍率で観察し，'
      'ディンプル，準へき開面，リバーパターン等の微細形態と'
      '第 1 章・第 3 章の力学的測定結果との対応を考察する．')

    h1(doc,'4.2　原理または理論')
    h2(doc,'4.2.1　走査型電子顕微鏡（SEM）の原理')
    B('タングステン陰極から射出した一次電子ビームを電子レンズで絞り，試料表面を走査する．'
      '照射された一次電子が試料と相互作用して放出される二次電子を'
      'シンチレーターおよび光電子増倍管で増幅・検出して表面形状像を得る 1)．'
      '観察倍率は数百倍から数万倍程度であり，破面の微細形態の観察に適する．'
      '試料は導電性が必要であり，絶縁体の場合は金蒸着等の導電処理を施す．')

    h2(doc,'4.2.2　延性・脆性破壊の破面形態')
    B('延性破壊ではボイドの発生・成長・合体を経た破断であり，破面にはディンプル（半球状窪み）が観察される 1)．'
      'SS400 はカップコーン破面を示し，中央の繊維状領域と外周のせん断リップ（45° 方向）が特徴的である 1)．'
      'カップコーン破面では，中央繊維状域（三軸引張応力場，等軸ディンプル）・遷移域・外周せん断リップ（平面応力，せん断ディンプル）の三領域が区別される 1)．'
      '脆性破壊では特定の結晶面に沿ったへき開破断が生じ，破面はへき開面（平坦光沢面）や'
      'リバーパターン（高さの異なるへき開面間の段差によって形成される河状模様）を示す 1)．'
      'FC250 ではフレーク黒鉛周辺から伝播したき裂による準へき開面・リバーパターンが形成される 1)4)．'
      'A7075 は高強度 Al 合金であるが，析出物（MgZn₂ 等）を核としたボイドコアレッセンスによる延性的破壊を示しつつ，'
      '一部でせん断成分が支配的な混合形態をとることが知られる 5)．')

    h1(doc,'4.3　装置の説明')
    fn_sem = fn()
    B(f'走査型電子顕微鏡（SEM，Fig. {fn_sem}）で破面観察を行った．')
    fig_caption(doc, fn_sem, '走査型電子顕微鏡 原理図（実験書参照）')

    h1(doc,'4.4　実験方法')
    B('ファインカッターで SS400，FC250，A7075 の引張試験後の破断部を切り出し，アセトンで洗浄・乾燥後，'
      'ビニール手袋を着用して試料台に固定した．'
      'SEM 内部を真空に排気後，各試料について×100 および ×500 の 2 倍率で'
      '中央繊維状域（c），カップ中央部（CC），遷移域（CN），外縁部（CO），せん断リップ（E）の'
      '5 部位を系統的に観察し，写真撮影を行った（計 30 枚）．')

    h1(doc,'4.5　観察結果')

    # --- Macro figures ---
    fn_ss400_frac = fn()
    fn_fc250_frac = fn()
    fn_a7075_frac = fn()
    B(f'各試験片の破断面マクロ写真・スケッチを Fig. {fn_ss400_frac}（SS400），'
      f'Fig. {fn_fc250_frac}（FC250），Fig. {fn_a7075_frac}（A7075）に貼付する．')
    fig_caption(doc, fn_ss400_frac, 'SS400 破断面マクロ（写真・スケッチ貼付）')
    fig_caption(doc, fn_fc250_frac, 'FC250 破断面マクロ（写真・スケッチ貼付）')
    fig_caption(doc, fn_a7075_frac, 'A7075 破断面マクロ（写真・スケッチ貼付）')
    B('マクロ観察では，SS400 では顕著なくびれが生じ，破面はカップコーン形状を呈した．'
      '外周部に 45° のせん断リップが認められ，中央部は繊維状であった．'
      'A7075 ではくびれは SS400 より小さく，破面は比較的平坦傾向であった．'
      'FC250 ではくびれはほとんど認められず，破面は平坦で光沢があり脆性破断の特徴を示した．')

    # ---- 4.5.1 SS400 ----
    h2(doc,'4.5.1　SS400 の破面観察')
    B('SS400 は典型的な延性破壊を呈し，カップコーン形状の破面全域にわたってディンプルが観察された．'
      '各部位の SEM 観察結果を以下に示す（左：×100，右：×500）．')

    regions_ss = [
        ('c',  '中央繊維状域',
         '×100 では細かく均一なディンプル模様が全面を覆い，典型的な延性破壊の外観を呈する（Fig. {fn}）．'
         '×500 では直径 10〜30 μm 程度の等軸ディンプルが明瞭に確認でき，'
         'ボイドコアレッセンスによる古典的な延性破面の特徴を示す．'),
        ('CC', 'カップ中央部',
         'カップ中央部（Fig. {fn}）は中央繊維状域と隣接する領域であり，×100 で粗い繊維状ネットワーク構造が観察される．'
         'ディンプルのサイズはやや大きく，延性的な破面性状が維持されている．'),
        ('CN', 'カップ-コーン遷移域',
         '×100 では上下 2 領域の明瞭な水平境界が確認できる（Fig. {fn}）．'
         '×500 では上側（コーン側）に平坦な板状形態（準へき開様）が，下側（カップ側）に細かな繊維状ゾーンが観察され，'
         'カップ-コーン境界での破壊機構の転換を示している．'),
        ('CO', 'カップ外縁部',
         '外縁部（Fig. {fn}）では破面に顕著な高さの段差が観察され，'
         '引張軸に対して 45° 方向に破壊方向が変化するカップ外縁特有の形態を示す．'),
        ('E',  'せん断リップ',
         'せん断リップ（Fig. {fn}）は最外周の平面応力支配域であり，'
         '×500 では細かく均一なせん断ディンプルが観察され，せん断変形によって変形した延性破面の特徴を示す．'),
    ]

    for region, region_name, desc_template in regions_ss:
        h3(doc, f'（{region_name}）')
        fn_pair = fn()
        desc = desc_template.replace('{fn}', str(fn_pair))
        B(desc)
        insert_sem_pair(doc,
                        sem('SS400', region, '100'),
                        sem('SS400', region, '500'))
        fig_caption(doc, fn_pair,
                    f'SS400 {REGION_JP[region]}（左：×100，右：×500）')

    # ---- 4.5.2 FC250 ----
    h2(doc,'4.5.2　FC250 の破面観察')
    B('FC250 は脆性破壊を呈し，フレーク黒鉛を起点とした準へき開面が破面全域で観察された．'
      '各部位の SEM 観察結果を以下に示す（左：×100，右：×500）．')

    regions_fc = [
        ('c',  '中央繊維状域',
         '×100 では粗い粒状テクスチャが全面を覆い，ディンプルは観察されない（Fig. {fn}）．'
         '×500 では平坦で角ばった板状形態（準へき開ファセット）が粒状バックグラウンドの中に明確に確認でき，'
         'フレーク黒鉛周辺からのき裂伝播による脆性破壊の特徴を示す．'),
        ('CC', 'カップ中央部',
         '中央脆性域の中央部（Fig. {fn}）では粗い粒状テクスチャに加えてより大きな準へき開ファセットが観察され，'
         '局所的なき裂伝播の痕跡が確認できる．'),
        ('CN', '中央脆性域',
         '中央脆性域（Fig. {fn}）ではほぼ平坦な外観を示し，起伏が少ない．'
         'フレーク黒鉛の方向に沿ったき裂伝播面が支配的であると推察される．'),
        ('CO', '外縁脆性域',
         '外縁脆性域（Fig. {fn}）は破面外周部に位置し，'
         'き裂が試験片表面に向かって伝播した痕跡が観察される．'),
        ('E',  'へき開域',
         'へき開域（Fig. {fn}）では角ばった準へき開ファセット（平坦板状小面）が明瞭に観察され，'
         'FC250 の脆性破壊に特徴的な形態である．'
         'ディンプルは認められず，全域が脆性的な破壊機構によって形成されたことを示す．'),
    ]

    for region, region_name, desc_template in regions_fc:
        h3(doc, f'（{region_name}）')
        fn_pair = fn()
        desc = desc_template.replace('{fn}', str(fn_pair))
        B(desc)
        insert_sem_pair(doc,
                        sem('FC250', region, '100'),
                        sem('FC250', region, '500'))
        fig_caption(doc, fn_pair,
                    f'FC250 {REGION_JP_FC[region]}（左：×100，右：×500）')

    # ---- 4.5.3 A7075 ----
    h2(doc,'4.5.3　A7075 の破面観察')
    B('A7075 は延性的破壊とせん断成分が混在する混合型の破面形態を示した．'
      '各部位の SEM 観察結果を以下に示す（左：×100，右：×500）．')

    regions_a7 = [
        ('c',  '中央繊維状域',
         '×100 では不規則な大きな塊状形態と繊維状特徴が混在して観察される（Fig. {fn}）．'
         '析出物（MgZn₂ 等）を核としたボイドコアレッセンスと，せん断成分による変形が複合した'
         'A7075 特有の混合破壊形態を示す．'),
        ('CC', 'カップ中央部',
         'カップ中央部（Fig. {fn}）では特徴的な縦方向への細長い条痕（ストライエーション様リッジ）が観察され，'
         'A7075 のせん断成分が支配的な部位に特有の縦長せん断ディンプルを示す．'),
        ('CN', 'カップ-コーン遷移域',
         '遷移域（Fig. {fn}）では中央部から外縁部への形態の変化が観察される．'
         '縦長のリッジ構造が徐々に変化し，破壊方向の転換を反映した形態を示す．'),
        ('CO', 'カップ外縁部',
         '外縁部（Fig. {fn}）ではカップ外周特有の高さ段差構造が観察され，'
         '破壊方向が 45° 方向へ転換する境界を反映した形態を示す．'),
        ('E',  'せん断リップ',
         'せん断リップ（Fig. {fn}）は比較的滑らかな外観を示しつつ，'
         '不規則な特徴も混在する．平面応力状態でのせん断変形が支配的であることを示す．'),
    ]

    for region, region_name, desc_template in regions_a7:
        h3(doc, f'（{region_name}）')
        fn_pair = fn()
        desc = desc_template.replace('{fn}', str(fn_pair))
        B(desc)
        insert_sem_pair(doc,
                        sem('A7075', region, '100'),
                        sem('A7075', region, '500'))
        fig_caption(doc, fn_pair,
                    f'A7075 {REGION_JP[region]}（左：×100，右：×500）')

    h1(doc,'4.6　課題')
    h3(doc,'課題 15　破断した各試験片の破面状態をスケッチせよ．')
    B(f'Fig. {fn_ss400_frac}〜Fig. {fn_a7075_frac} に各試験片の破断面マクロを示す．'
      'SS400 はカップコーン形状（中央繊維状部 + 外周 45° せん断リップ）を呈した延性破断であった．'
      'A7075 は SS400 より小さいくびれを伴う延性的破断，FC250 は平坦光沢面の脆性破断であった．'
      'SEM 観察（4.5 節）では各部位の微細形態が確認された．')

    h3(doc,'課題 16　同一観察場所において，材料の違いによる破面の違いについて考察せよ．')
    B('中央繊維状域（c 領域）における材料間の比較を SEM 観察結果に基づいて述べる．')
    B('SS400（延性）：×500 で直径 10〜30 μm の等軸ディンプルが密集して観察された（4.5.1 節）．'
      'ボイドコアレッセンス機構による典型的な延性破面であり，三軸引張応力場のもとで'
      'ボイドが等方的に成長・合体した結果である 1)．')
    B('FC250（脆性）：×500 で平坦な板状の準へき開ファセットが粒状バックグラウンドの中に観察された（4.5.2 節）．'
      'ディンプルは認められず，フレーク黒鉛の先端を起点としたき裂が特定の結晶面に沿って伝播した'
      '脆性的破壊の特徴を示す 1)4)．')
    B('A7075（混合）：×100 で大きな塊状形態と繊維状特徴が混在して観察された（4.5.3 節）．'
      'MgZn₂ 等の析出物を核としたボイドコアレッセンスと局所的なせん断変形が複合した形態であり，'
      'SS400 ほど均一なディンプルは形成されない 5)．')
    B('以上より，同一部位における破面形態の差異は，延性・脆性の序列（SS400 >> A7075 > FC250）と'
      '整合しており，第 1 章の引張試験で得た破断伸び δ の序列（28.58% >> 7.69% > 4.84%）を'
      '微細構造レベルで直接的に裏付けている．')

    h3(doc,'課題 17　同一材料において，観察場所による破面の違いについて考察せよ．')
    B('SS400 のカップコーン破面を例に，SEM 観察結果（4.5.1 節）に基づいて各部位の形態の違いを述べる．')
    B('（a）中央繊維状域（c 領域）：試験片軸に垂直な内部領域であり，'
      '三軸引張応力場のもとでボイドが発生・成長・合体する延性的破壊が支配的な部位である 1)．'
      'SEM 観察では直径 10〜30 μm の細かく均一な等軸ディンプルが観察された．')
    B('（b）カップ中央部（CC 領域）・遷移域（CN 領域）：カップ中央部ではやや大きなディンプルが観察され，'
      '遷移域では×100 で明瞭な水平境界が，×500 でコーン側の準へき開様形態とカップ側の繊維状ゾーンの共存が確認された．'
      'この境界は平面ひずみ（内部）から平面応力（外縁）への応力状態の変化を反映している 1)．')
    B('（c）外縁部（CO 領域）：顕著な高さ段差が観察され，引張軸に対して 45° 方向への破壊方向の転換を示す 1)．')
    B('（d）せん断リップ（E 領域）：最外周の平面応力域であり，細かく均一なせん断ディンプルが観察された．'
      'せん断応力支配によりディンプルの形状は等軸ではなく，変形方向に引き伸ばされた形態をとる場合がある 1)．')

    h1(doc,'4.7　考察')
    B(f'各材料の SEM 観察結果は，第 1 章の引張試験で得た延性特性と整合した 1)．'
      f'SS400（δ = {KV["SS400"]["delta"]:.2f}%）の破面全域にわたるディンプルは，延性的なボイドコアレッセンス機構の直接的証拠である．'
      f'FC250（δ = {KV["FC250"]["delta"]:.2f}%）の準へき開ファセットは，フレーク黒鉛を起点とした脆性的なき裂伝播に対応する 4)．'
      f'A7075（δ = {KV["A7075"]["delta"]:.2f}%）の混合形態は，高強度 Al 合金に特有の析出物核ボイドコアレッセンスとせん断変形の複合機構を反映している 5)．')
    B('SS400 のカップコーン破面における場所依存性（中央繊維状域での等軸ディンプル → 遷移域での二相共存 → せん断リップでのせん断ディンプル）は，'
      '試験片内部（三軸引張応力）から表面近傍（平面応力）への応力状態の変化に起因する 1)．'
      '遷移域における明瞭な水平境界（×100 観察）とその内部の準へき開様形態（×500 観察）は，'
      '破壊先進と遅行の空間的分布を視覚的に示すものである．')
    B('FC250 の場合，中央部・外縁部いずれも延性的形態は観察されず，全域が脆性的な準へき開形態を示した．'
      'これはフレーク黒鉛の均一分散が試験片全断面にわたって応力集中を生じさせ，'
      '延性的破壊プロセスを完全に抑制していることを示す 4)．')

    # ===== 感想 =====
    doc.add_page_break()
    h1(doc,'感想')
    para(doc,'',indent_mm=0)
    B('本実験を通じて，SS400，FC250，A7075 という 3 種類の材料が，'
      '同じ引張試験に対してまったく異なる応答を示すことを数値として直接確認できた．'
      f'特に FC250 の低延性（{KV["FC250"]["delta"]:.2f}%）と SS400 の高延性（{KV["SS400"]["delta"]:.2f}%）の約 6 倍の差は，'
      '材料組織（フレーク黒鉛の形状・分布）が力学特性に直接影響することを実感させるものであった．')
    B('追加実験では LVDT 伸び計の有無によるヤング率の差から，試験機コンプライアンスが'
      '定量的に評価できることを学んだ．正確な材料特性を得るためには計測系の適切な選択が重要であると認識した．')
    B('シャルピー試験では振り子型ハンマの落下という単純な操作から吸収エネルギーを定量化できる点に'
      '試験設計の工夫を感じた．beta 角の精度が結果に直接影響するため，角度読取りの重要性を改めて認識した．')
    B('破面観察では SEM を用いてミクロな破壊機構を直接観察できる点が印象的であった．'
      'SS400 のカップコーン破面と FC250 の平坦破面という，巨視的特性とミクロ形態の対応が明確に確認できた 1)4)．')

    # ===== 備考 =====
    h1(doc,'備考')
    para(doc,'',indent_mm=0)
    B(f'Fig. 2 は SS400 Hollomon 則の両対数フィット（実測データ，K = 908 MPa，n = 0.292）である．'
      f'Fig. 3〜6 は実測データ（SS400: 41 点，A7075: 9 点）に基づく線図である．'
      'FC250 の応力-ひずみ線図は参考データ（CSV 上 epsilon_max = 1.08%）であり，'
      '実測 delta = 4.84% との整合性については実験書原本データを要確認とする．')
    B('Fig. 7（SS400 通常 vs 追加実験比較）は追加実験の伸び計データが未確定のため概略図のままである．'
      '伸び計データが得られた場合は実測点で更新する．')
    B('Table 3 の絞り phi 欄は体積一定仮定に基づく近似値であり，'
      '正確な値は破断部最小径（マイクロメーター計測）から算出する．'
      'Table 3 の E 実験欄および Table 6 の K・rho 欄は，以下の実測値が確定次第追記する．'
      '（a）beta 角（シャルピー読取り値） → K および rho の算出'
      '（b）弾性域グラフデータ（伸び計） → E 実験値の算出')
    B('有効数字については，直径測定値（0.001 mm 精度）は 5 桁，'
      '標点間距離（0.01 mm 精度）は 4〜5 桁で扱った．'
      '最終的な機械的性質の有効数字は実験精度全体を考慮して 3 桁程度が適切と考えられる．')

    # ===== 参考文献 =====
    doc.add_page_break()
    h1(doc,'参考文献')
    para(doc,'',indent_mm=0)
    for ref in [
        '1) 小寺沢良一，フラクトグラフィ，培風館 (1985)，pp. 1-50．',
        '2) 日本機械学会編，機械工学便覧 基礎編 alpha1 材料力学，日本機械学会 (2004)，pp. 1-45．',
        '3) 日本工業標準調査会，JIS Z 2242 金属材料のシャルピー衝撃試験方法，日本規格協会 (2018)．',
        '4) 日本機械学会編，機械材料学，日本機械学会 (2007)，pp. 120-140．',
        '5) 軽金属学会編，アルミニウムの組織と性質，軽金属学会 (1991)，pp. 201-215．',
        '6) 日本機械学会編，機械工学便覧 基礎編 alpha4 計測工学，日本機械学会 (2007)，pp. 4-25．',
    ]:
        para(doc,ref,size=10,sa=2,indent_mm=0)

    # ===== Unicode 添字変換（Windows互換処理）=====
    fix_subscripts_in_doc(doc)

    return doc

print('[v7] ビルド開始...')
doc = build()
doc.save(OUTPUT_PATH)
print(f'[v7] 完了: {OUTPUT_PATH}')
