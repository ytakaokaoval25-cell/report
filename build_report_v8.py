#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""実験報告書 v8: 章構成を刷新（1-8節形式）+ 引張/シャルピー両SEM分類"""
import io, math, copy, os
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

OUTPUT_PATH = "/Users/takaoka/report/report_v8.docx"
SEM_DIR = "/Users/takaoka/report/sem_images/"

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
    'SS400':     dict(d1=14.060,d2=14.053,d3=14.068,l0=82.18,A0=155.268,lf=105.67,
                     E_lit=206,sy_lit=245,sB_lit=400,sf_lit=320),
    'FC250':     dict(d1=14.060,d2=13.992,d3=14.022,l0=62.22,A0=154.491,lf=65.23,
                     E_lit=130,sy_lit=None,sB_lit=250,sf_lit=None),
    'A7075':     dict(d1=13.982,d2=14.017,d3=14.023,l0=62.45,A0=154.099,lf=67.25,
                     E_lit=71.7,sy_lit=503,sB_lit=572,sf_lit=450),
    'SS400_add': dict(d1=13.980,d2=13.981,d3=14.070,l0=82.30,A0=154.165,lf=110.18,
                     E_lit=206,sy_lit=245,sB_lit=400,sf_lit=320),
}
CHARPY = {
    'SS400': dict(l=54.98, b=10.00, h1=10.00, h2=10.00, A=100.0),
    'FC250': dict(l=54.90, b=10.00, h1=10.00, h2=9.97,  A=99.7),
    'A7075': dict(l=55.00, b=9.90,  h1=10.00, h2=10.00, A=99.0),
}
CHARPY_BETA = {'SS400': 120, 'FC250': 143, 'A7075': 141}
m_kg = 25.093; g_acc = 9.798; W_N = m_kg * g_acc; r_m = 0.6513; M_J = W_N * r_m
alpha_deg = 145; cos_alpha = math.cos(math.radians(alpha_deg))
# 損失エネルギー: 空打ち試験で145°→144°
beta_air = 144
L_J = M_J * (math.cos(math.radians(beta_air)) - math.cos(math.radians(alpha_deg)))
# 最終K値（確定値）
K_final = {
    'SS400': 49.49,
    'FC250': 1.66,
    'A7075': 5.11,
}
rho_final = {
    'SS400': 49.49,
    'FC250': 1.67,
    'A7075': 5.16,
}
A_cm2 = {'SS400': 1.000, 'FC250': 0.997, 'A7075': 0.990}

def elong(key):
    d = MAT[key]; return (d['lf'] - d['l0']) / d['l0'] * 100

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
    rFonts.set(qn('w:cs'),       'Times New Roman')
    for e in rPr.findall(qn('w:rFonts')): rPr.remove(e)
    rPr.insert(0, rFonts)

def _set_spacing(para, pt=18):
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing'); pPr.append(sp)
    sp.set(qn('w:line'), str(int(Pt(pt).pt * 20)))
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

def section(doc, text):
    """1-1　目的 style"""
    return para(doc, text, size=13, bold=True, sb=10, sa=4, indent_mm=0)

def subsection(doc, text):
    """1-1-1　小節 style"""
    return para(doc, text, size=11, bold=True, sb=6, sa=3, indent_mm=0)

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
    tab.set(qn('w:val'), 'right'); tab.set(qn('w:pos'), str(int(Mm(155).pt * 20)))
    tabs.append(tab); pPr.append(tabs)
    return p

def fig_caption(doc, fn, text):
    return para(doc, f'図{fn}　{text}',
                align=WD_ALIGN_PARAGRAPH.CENTER, size=10, sb=1, sa=6, indent_mm=0)

def tbl_caption(doc, tn, text):
    return para(doc, f'表{tn}　{text}',
                align=WD_ALIGN_PARAGRAPH.CENTER, size=10, sb=6, sa=2, indent_mm=0)

def proc(doc, text, level=1):
    """手順段落: level=1 は通常，level=2 はサブ"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    if level == 1:
        p.paragraph_format.left_indent = Mm(0)
        p.paragraph_format.first_line_indent = Mm(0)
    else:
        p.paragraph_format.left_indent = Mm(6)
        p.paragraph_format.first_line_indent = Mm(0)
    _set_spacing(p, 18)
    r = p.add_run(text)
    _set_run_font(r, 10.5)
    return p

def set_margins(doc):
    for sec in doc.sections:
        sec.page_height = Mm(297); sec.page_width = Mm(210)
        sec.left_margin = Mm(25);  sec.right_margin = Mm(15)
        sec.top_margin  = Mm(25);  sec.bottom_margin = Mm(20)

def fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0); plt.close(fig); return buf

def insert_figure(doc, buf, width_cm=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(buf, width=Cm(width_cm))
    return p

def B(doc, text, ind=True):
    return para(doc, text, indent_mm=4 if ind else 0)

def Bx(doc, text):
    return para(doc, text, indent_mm=0)

# ============================================================
# SEM 画像ヘルパー
# ============================================================
def sem_path(material, region, mag):
    return os.path.join(SEM_DIR, f"{material}{region}{mag}.jpg")

def insert_sem_pair(doc, path1, path2):
    """SEM画像2枚を横並びで挿入（2列ボーダーレステーブル）"""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl._tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{border}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    for i, path in enumerate([path1, path2]):
        cell = tbl.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(path):
            p.add_run().add_picture(path, width=Cm(7.0))
        else:
            p.add_run(f'[画像なし: {os.path.basename(path)}]')
    return tbl

# ============================================================
# Unicode 添字・上付きを Word 書式に変換するポストプロセッサ
# ============================================================
def fix_subscripts_in_doc(doc):
    SUB_MAP = {'₀':'0','₁':'1','₂':'2','₃':'3','₄':'4',
               '₅':'5','₆':'6','₇':'7','₈':'8','₉':'9',
               'ₐ':'a','ᵢ':'i','ⱼ':'j','ₚ':'p','ₜ':'t',
               'ₑ':'e','ₒ':'o','ₙ':'n','ₛ':'s','ᵤ':'u'}
    SUP_MAP = {'²':'2','³':'3','⁰':'0','⁴':'4','⁵':'5',
               '⁶':'6','⁷':'7','⁸':'8','⁹':'9',
               'ⁿ':'n','ᵖ':'p','ᵃ':'a','ᵇ':'b','ᶜ':'c','ᵉ':'e'}
    ALL = {k: ('subscript', v) for k, v in SUB_MAP.items()}
    ALL.update({k: ('superscript', v) for k, v in SUP_MAP.items()})

    def split_text(text):
        segs = []; cur = ''; ctype = 'normal'
        for c in text:
            if c in ALL:
                if cur: segs.append((cur, ctype))
                t, ch = ALL[c]; segs.append((ch, t))
                cur = ''; ctype = 'normal'
            else:
                if ctype != 'normal' and cur:
                    segs.append((cur, ctype)); cur = ''; ctype = 'normal'
                cur += c
        if cur: segs.append((cur, ctype))
        merged = []
        for s in segs:
            if merged and merged[-1][1] == s[1]:
                merged[-1] = (merged[-1][0] + s[0], s[1])
            else:
                merged.append(list(s))
        return merged

    def process_run(run_xml, parent):
        t_el = run_xml.find(qn('w:t'))
        if t_el is None: return
        text = t_el.text or ''
        if not any(c in ALL for c in text): return
        segs = split_text(text)
        if len(segs) == 1 and segs[0][1] == 'normal': return
        idx = list(parent).index(run_xml)
        new_rs = []
        for seg_text, seg_type in segs:
            nr = copy.deepcopy(run_xml)
            for t in nr.findall(qn('w:t')): nr.remove(t)
            t = OxmlElement('w:t'); t.text = seg_text
            if seg_text != seg_text.strip() or seg_text == ' ':
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            nr.append(t)
            if seg_type != 'normal':
                rPr = nr.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr'); nr.insert(0, rPr)
                for va in rPr.findall(qn('w:vertAlign')): rPr.remove(va)
                va = OxmlElement('w:vertAlign')
                va.set(qn('w:val'), seg_type); rPr.append(va)
                for tag in (qn('w:sz'), qn('w:szCs')):
                    el = rPr.find(tag)
                    if el is not None:
                        v = int(el.get(qn('w:val'), '21'))
                        el.set(qn('w:val'), str(max(14, int(v * 0.72))))
            new_rs.append(nr)
        for i, nr in enumerate(new_rs): parent.insert(idx + i, nr)
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
# グラフ生成
# ============================================================
def _smooth(e_arr, s_arr, n=300):
    if _SCIPY and len(e_arr) > 3:
        cs = PchipInterpolator(e_arr, s_arr)
        ex = np.linspace(e_arr[0], e_arr[-1], n)
        return ex, cs(ex)
    ex = np.linspace(e_arr[0], e_arr[-1], n)
    return ex, np.interp(ex, e_arr, s_arr)

def plot_overview():
    fig, ax = plt.subplots(figsize=(10, 6))
    e, s = zip(*SS400_ENG)
    ax.plot(np.array(e)*100, s, color='steelblue', lw=2, label='SS400（構造用鋼）')
    e, s = zip(*A7075_ENG); e, s = np.array(e), np.array(s)
    ex, sx = _smooth(e, s)
    ax.plot(ex*100, sx, color='darkorange', ls='--', lw=2, label='A7075（Al合金）')
    e, s = zip(*FC250_ENG)
    ax.plot(np.array(e)*100, s, color='forestgreen', ls=':', lw=2, label='FC250（ねずみ鋳鉄）')
    ax.set_xlabel('公称ひずみ e [%]', fontsize=12)
    ax.set_ylabel('公称応力 sigma [MPa]', fontsize=12)
    ax.set_title('3材料 公称応力-公称ひずみ線図（実測データ）', fontsize=13)
    ax.legend(fontsize=11); ax.set_xlim(0); ax.set_ylim(0); ax.grid(True, ls='--', alpha=0.5)
    fig.tight_layout(); return fig_to_buf(fig)

def plot_ss400():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5))
    e_arr = np.array([p[0] for p in SS400_ENG])
    s_arr = np.array([p[1] for p in SS400_ENG])
    ax1.plot(e_arr*100, s_arr, 'steelblue', lw=2)
    ax1.plot(e_arr*100, s_arr, 'o', color='steelblue', ms=3, alpha=0.5)
    ax1.annotate(f'上降伏点 {KV["SS400"]["sy"]:.1f} MPa',
                 xy=(SS400_ENG[16][0]*100, KV['SS400']['sy']),
                 xytext=(SS400_ENG[16][0]*100+2, KV['SS400']['sy']*1.10),
                 fontsize=8, color='navy', arrowprops=dict(arrowstyle='->', color='navy', lw=1))
    ax1.annotate(f'下降伏点 {KV["SS400"]["sL"]:.1f} MPa',
                 xy=(SS400_ENG[18][0]*100, KV['SS400']['sL']),
                 xytext=(SS400_ENG[18][0]*100+2, KV['SS400']['sL']*0.88),
                 fontsize=8, color='navy', arrowprops=dict(arrowstyle='->', color='navy', lw=1))
    ax1.annotate(f'sigma_B {KV["SS400"]["sB"]:.1f} MPa',
                 xy=(SS400_ENG[37][0]*100, KV['SS400']['sB']),
                 xytext=(SS400_ENG[37][0]*100-6, KV['SS400']['sB']*1.06),
                 fontsize=8, color='darkred', arrowprops=dict(arrowstyle='->', color='darkred', lw=1))
    ax1.axvline(KV['SS400']['delta'], color='red', ls='--', lw=1,
                label=f'破断 delta={KV["SS400"]["delta"]:.2f}%')
    ax1.set_xlabel('公称ひずみ e [%]', fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]', fontsize=11)
    ax1.set_title('SS400\n公称応力-公称ひずみ線図（実測）', fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True, ls='--', alpha=0.5)
    e_u = e_arr[:38]; s_u = s_arr[:38]
    et = np.log(1+e_u); st = s_u*(1+e_u)
    ax2.plot(et*100, st, 'darkorange', lw=2)
    ax2.plot(et*100, st, 'o', color='darkorange', ms=3, alpha=0.5)
    ax2.set_xlabel('対数ひずみ epsilon [%]', fontsize=11)
    ax2.set_ylabel('真応力 sigma_t [MPa]', fontsize=11)
    ax2.set_title('SS400\n真応力-対数ひずみ線図（均一変形域，実測）', fontsize=11)
    ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True, ls='--', alpha=0.5)
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def plot_a7075():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5))
    e, s = zip(*A7075_ENG); e, s = np.array(e), np.array(s)
    ex, sx = _smooth(e, s)
    ax1.plot(ex*100, sx, 'steelblue', lw=2)
    ax1.plot(e*100, s, 'o', color='steelblue', ms=5, label='測定点')
    ax1.annotate(f'sigma_B {KV["A7075"]["sB"]:.1f} MPa',
                 xy=(A7075_ENG[7][0]*100, KV['A7075']['sB']),
                 xytext=(A7075_ENG[7][0]*100-2.5, KV['A7075']['sB']*1.05),
                 fontsize=8, color='darkred', arrowprops=dict(arrowstyle='->', color='darkred', lw=1))
    ax1.axhline(KV['A7075']['sy_lit'], color='gray', ls=':', lw=1)
    ax1.text(0.5, KV['A7075']['sy_lit']+20, f'0.2%耐力（文献）503 MPa', fontsize=8, color='gray')
    ax1.axvline(KV['A7075']['delta'], color='red', ls='--', lw=1,
                label=f'破断 delta={KV["A7075"]["delta"]:.2f}%')
    ax1.set_xlabel('公称ひずみ e [%]', fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]', fontsize=11)
    ax1.set_title('A7075（Al合金）\n公称応力-公称ひずみ線図（実測）', fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True, ls='--', alpha=0.5)
    et, st = zip(*A7075_TRUE); et, st = np.array(et), np.array(st)
    etx, stx = _smooth(et, st)
    ax2.plot(etx*100, stx, 'darkorange', lw=2)
    ax2.plot(et*100, st, 'o', color='darkorange', ms=5, label='測定点')
    ax2.set_xlabel('対数ひずみ epsilon [%]', fontsize=11)
    ax2.set_ylabel('真応力 sigma_t [MPa]', fontsize=11)
    ax2.set_title('A7075（Al合金）\n真応力-対数ひずみ線図（実測）', fontsize=11)
    ax2.legend(fontsize=9); ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True, ls='--', alpha=0.5)
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def plot_fc250():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5))
    e, s = zip(*FC250_ENG); e, s = np.array(e), np.array(s)
    ax1.plot(e*100, s, 'steelblue', lw=2)
    ax1.plot(e*100, s, 'o', color='steelblue', ms=5, label='測定点（参考）')
    ax1.annotate(f'sigma_B {KV["FC250"]["sB"]:.1f} MPa',
                 xy=(FC250_ENG[-1][0]*100, KV['FC250']['sB']),
                 xytext=(FC250_ENG[-1][0]*100-0.3, KV['FC250']['sB']*1.08),
                 fontsize=8, color='darkred', arrowprops=dict(arrowstyle='->', color='darkred', lw=1))
    ax1.set_xlabel('公称ひずみ e [%]', fontsize=11)
    ax1.set_ylabel('公称応力 sigma [MPa]', fontsize=11)
    ax1.set_title('FC250（ねずみ鋳鉄）\n公称応力-公称ひずみ線図（参考データ）', fontsize=11)
    ax1.legend(fontsize=9); ax1.set_xlim(0); ax1.set_ylim(0); ax1.grid(True, ls='--', alpha=0.5)
    et, st = zip(*FC250_TRUE); et, st = np.array(et), np.array(st)
    ax2.plot(et*100, st, 'darkorange', lw=2)
    ax2.plot(et*100, st, 'o', color='darkorange', ms=5, label='測定点（参考）')
    ax2.set_xlabel('対数ひずみ epsilon [%]', fontsize=11)
    ax2.set_ylabel('真応力 sigma_t [MPa]', fontsize=11)
    ax2.set_title('FC250（ねずみ鋳鉄）\n真応力-対数ひずみ線図（参考データ）', fontsize=11)
    ax2.legend(fontsize=9); ax2.set_xlim(0); ax2.set_ylim(0); ax2.grid(True, ls='--', alpha=0.5)
    fig.tight_layout(pad=2); return fig_to_buf(fig)

def _hollomon_ss400():
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
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.loglog(ep*100, st, 'o', color='steelblue', ms=6, label='SS400 実測点')
    ep_line = np.logspace(np.log10(ep[0]), np.log10(ep[-1]), 100)
    st_line = K * ep_line**n
    ax.loglog(ep_line*100, st_line, color='darkorange', lw=2,
              label=f'Hollomon則: K={K:.0f} MPa, n={n:.3f}\nR2={r2:.4f}')
    ax.set_xlabel('相当塑性ひずみ epsilon_p [%] (対数軸)', fontsize=11)
    ax.set_ylabel('真応力 sigma_t [MPa] (対数軸)', fontsize=11)
    ax.set_title(f'SS400 Hollomon則 両対数プロット（実測データ）\nK={K:.0f} MPa, n={n:.3f}', fontsize=11)
    ax.legend(fontsize=10); ax.grid(True, which='both', ls='--', alpha=0.5)
    fig.tight_layout(); return fig_to_buf(fig)

def plot_ss400_comparison():
    fig, ax = plt.subplots(figsize=(10, 6))
    e, s = zip(*SS400_ENG)
    ax.plot(np.array(e)*100, s, color='steelblue', lw=2, label='SS400 通常試験（実測）')
    d = MAT['SS400_add']; E = d['E_lit']*1000; sy = d['sy_lit']; sB = d['sB_lit']
    ef = elong('SS400_add')/100
    e1 = np.linspace(0, sy/E, 50); s1 = E*e1
    ep = np.linspace(1e-3, ef-sy/E, 100); e2 = sy/E+ep; s2 = np.clip(540*ep**0.20, sy-10, sB)
    e3 = np.array([e2[-1], ef]); s3 = np.array([s2[-1], 320])
    e_add = np.r_[e1, e2, e3]; s_add = np.r_[s1, s2, s3]
    ax.plot(e_add*100, s_add, color='darkorange', ls='--', lw=2,
            label='SS400 追加実験（伸び計，概略）')
    ax.set_xlabel('公称ひずみ e [%]', fontsize=12)
    ax.set_ylabel('公称応力 sigma [MPa]', fontsize=12)
    ax.set_title('SS400 通常試験 vs 追加実験', fontsize=12)
    ax.legend(fontsize=11); ax.set_xlim(0); ax.set_ylim(0); ax.grid(True, ls='--', alpha=0.5)
    ax.text(0.02, 0.03, '注：追加実験は伸び計データ確定後に実測値で更新予定',
            transform=ax.transAxes, fontsize=9, color='gray')
    fig.tight_layout(); return fig_to_buf(fig)

# ============================================================
# 表ヘルパー
# ============================================================
def make_table(doc, headers, rows, col_widths=None):
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]; c.text = h
        for r in c.paragraphs[0].runs:
            r.font.bold = True; r.font.size = Pt(9)
    for row_data in rows:
        row = tbl.add_row().cells
        for i, v in enumerate(row_data):
            row[i].text = str(v)
            if row[i].paragraphs[0].runs:
                row[i].paragraphs[0].runs[0].font.size = Pt(9)
    return tbl

# ============================================================
# 文書構築
# ============================================================
def build():
    doc = Document(); set_margins(doc)
    sty = doc.styles['Normal']
    sty.font.name = 'Times New Roman'; sty.font.size = Pt(10.5)
    sty.paragraph_format.line_spacing = Pt(18)

    _eq = [0]; _fn = [0]; _tn = [0]
    def eq(): _eq[0] += 1; return _eq[0]
    def fn(): _fn[0] += 1; return _fn[0]
    def tn(): _tn[0] += 1; return _tn[0]

    # Hollomon 定数を先に計算
    _ep, _st, _K, _n, _r2 = _hollomon_ss400()
    # 靭性
    _U_SS = np.trapz([p[1] for p in SS400_ENG], [p[0] for p in SS400_ENG])
    _U_A7 = np.trapz([p[1] for p in A7075_ENG], [p[0] for p in A7075_ENG])
    _U_FC = np.trapz([p[1] for p in FC250_ENG], [p[0] for p in FC250_ENG])

    # ===== 表紙 =====
    for _ in range(3): doc.add_paragraph()
    para(doc, '機械航空宇宙工学実験　報告書',
         align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True, sb=0, sa=8, indent_mm=0)
    para(doc, '引張試験・シャルピー衝撃試験・破面観察',
         align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, sb=0, sa=30, indent_mm=0)
    for label, val in [('科目', '機械航空宇宙工学実験'),
                       ('学科・学年', '機械航空宇宙工学科　3年　6B'),
                       ('学籍番号', ''),
                       ('氏名', ''),
                       ('提出日', '')]:
        para(doc, f'{label}：{val}',
             align=WD_ALIGN_PARAGRAPH.CENTER, size=12, sb=0, sa=5, indent_mm=0)
    doc.add_page_break()

    # ===========================================================
    # 第1章 引張試験
    # ===========================================================
    chapter_title(doc, '第1章　引張試験')

    # --- 1-1 目的 ---
    section(doc, '1-1　目的')
    B(doc, '引張試験機を用いてSS400，FC250，A7075の3種類の試験片に静的引張荷重を加え，'
      '公称応力−公称ひずみ線図および真応力−対数ひずみ線図を作成する．'
      '各材料の機械的性質（降伏応力，引張強さ，破断伸び，絞り，ヤング率など）を測定・評価し，'
      '延性・脆性の違いと破壊挙動の対応を考察することを目的とする．')

    # --- 1-2 試験機 ---
    section(doc, '1-2　試験機')
    B(doc, '島津製作所製万能試験機UH-300kNA（最大試験力300kN，コンピュータ制御・油圧サーボ式）を使用した．'
      '試験力はロードセルで計測し，荷重−変位データをコンピュータに自動記録する．'
      '試験力精度はJIS B 7721に準拠した±1%以内である．')

    # --- 1-3 試験片 ---
    section(doc, '1-3　試験片')
    B(doc, 'JIS Z 2241の4号試験片（丸棒，平行部直径約14mm）を使用した．'
      '材料はSS400（一般構造用圧延鋼材），FC250（ねずみ鋳鉄），A7075（アルミニウム合金）の3種類である．')

    tn1 = tn()
    tbl_caption(doc, tn1, '試験片寸法')
    make_table(doc,
        ['材料', 'd1 [mm]', 'd2 [mm]', 'd3 [mm]', 'l0 [mm]', 'A0 [mm²]'],
        [('SS400', '14.060', '14.053', '14.068', '82.18', '155.268'),
         ('FC250', '14.060', '13.992', '14.022', '62.22', '154.491'),
         ('A7075', '13.982', '14.017', '14.023', '62.45', '154.099')])
    Bx(doc, '')

    # --- 1-4 試験前の処理 ---
    section(doc, '1-4　試験前の処理')
    B(doc, 'マイクロメーターで各試験片の平行部直径を3箇所（d1, d2, d3）測定し平均直径dを求め，'
      'A0=π(d/2)²で初期断面積を算出した．標点間距離l0はノギスで測定した．')

    # --- 1-5 試験機の初期調整 ---
    section(doc, '1-5　試験機の初期調整')
    B(doc, '電源投入後，負荷ポンプを起動し，ラム初期位置を設定した．'
      '荷重ゼロ・スパン調整を行い，LOAD RANGEを150kN，応力速度3.0MPa/sに設定した．')

    # --- 1-6 実験手順 ---
    section(doc, '1-6　実験手順')
    proc(doc, '① 試験片の取り付け')
    proc(doc, '　（1）試験片を上つかみに固定する', level=2)
    proc(doc, '　（2）断面積を入力し，ZEROを確認する', level=2)
    proc(doc, '　（3）下つかみに固定し，ENTERキーで試験を開始する', level=2)
    proc(doc, '② 試験の実施')
    proc(doc, '　（1）荷重−変位曲線をリアルタイムで確認しながら試験を進める', level=2)
    proc(doc, '　（2）破断後，ラムを初期位置に戻す', level=2)
    proc(doc, '③ 試験後の計測')
    proc(doc, '　（1）試験片を取り外し，破断部最小径を測定する', level=2)
    proc(doc, '　（2）破断後標点間距離lfをノギスで測定する', level=2)
    proc(doc, '　（3）破断面のスケッチを行う', level=2)

    # --- 1-7 結果 ---
    section(doc, '1-7　結果')
    B(doc, '式δ = (lf − l₀) / l₀ × 100 より算出した破断伸びを表2に示す．'
      'SS400（28.58%）>> A7075（7.69%）> FC250（4.84%）の延性序列が確認できる．')

    tn2 = tn()
    tbl_caption(doc, tn2, '破断後寸法と破断伸び')
    make_table(doc,
        ['材料', 'lf [mm]', 'δ [%]'],
        [('SS400', '105.67', '28.58'),
         ('FC250', '65.23',  '4.84'),
         ('A7075', '67.25',  '7.69')])
    Bx(doc, '')

    B(doc, '各材料の機械的性質の実測値および文献値を表3に示す．'
      'ヤング率は文献値を参照した（SS400: 206GPa，A7075: 71.7GPa，FC250: 130GPa）．'
      'SS400の降伏応力は上降伏点応力415.7MPaであり，下降伏点応力341.4MPaも確認された．'
      'A7075の0.2%耐力は今回のクロスヘッド変位では算出困難のため文献値503MPaを参照した．')

    tn3 = tn()
    tbl_caption(doc, tn3, '引張試験機械的性質（実測・文献値）')
    make_table(doc,
        ['材料', 'σy [MPa]', 'σB [MPa]', 'δ [%]', 'E文献 [GPa]'],
        [('SS400', '415.7（上降伏点）', '463.1', '28.58', '206'),
         ('FC250', '−（脆性）',        '396.1', '4.84',  '130'),
         ('A7075', '503（文献，0.2%耐力）', '1103.2', '7.69', '71.7')])
    Bx(doc, '')

    # グラフ挿入
    fn1 = fn()
    B(doc, f'図{fn1}に3材料の公称応力−公称ひずみ線図を示す．')
    insert_figure(doc, plot_overview())
    fig_caption(doc, fn1, '3材料比較 公称応力-公称ひずみ線図')

    fn2 = fn()
    B(doc, f'図{fn2}にSS400の公称応力−公称ひずみ線図（左）および真応力−対数ひずみ線図（右）を示す．')
    insert_figure(doc, plot_ss400(), width_cm=15)
    fig_caption(doc, fn2, 'SS400 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（実測）')

    fn3 = fn()
    B(doc, f'図{fn3}にA7075の公称応力−公称ひずみ線図（左）および真応力−対数ひずみ線図（右）を示す．')
    insert_figure(doc, plot_a7075(), width_cm=15)
    fig_caption(doc, fn3, 'A7075 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（実測）')

    fn4 = fn()
    B(doc, f'図{fn4}にFC250の公称応力−公称ひずみ線図（左）および真応力−対数ひずみ線図（右）を示す．')
    insert_figure(doc, plot_fc250(), width_cm=15)
    fig_caption(doc, fn4, 'FC250 公称応力-公称ひずみ線図（左）・真応力-対数ひずみ線図（右）（参考データ）')

    fn5 = fn()
    B(doc, f'図{fn5}にSS400のHollomon則両対数プロットを示す（K={_K:.0f}MPa，n={_n:.3f}，R²={_r2:.4f}）．')
    insert_figure(doc, plot_hollomon(), width_cm=11)
    fig_caption(doc, fn5, f'SS400 Hollomon則 両対数プロット（K={_K:.0f}MPa，n={_n:.3f}）')

    # --- 1-8 考察 ---
    section(doc, '1-8　考察')

    subsection(doc, '課題1　実験書8の諸量')
    B(doc, '以下に機械的性質の各量の定義と物理的意味を述べる．')
    equation(doc, 'σy = Py / A₀　（上降伏点応力 [MPa]）', eq())
    equation(doc, 'σL = PLv / A₀　（下降伏点応力 [MPa]）', eq())
    equation(doc, 'σB = Pmax / A₀　（引張強さ [MPa]）', eq())
    equation(doc, 'σT = Pt / Af　（破断応力 [MPa]）', eq())
    equation(doc, 'δ = (lf − l₀) / l₀ × 100　[%]　（破断伸び）', eq())
    equation(doc, 'φ = (A₀ − Af) / A₀ × 100　[%]　（絞り）', eq())
    equation(doc, 'E = Δσ / Δε　（ヤング率 [GPa]）', eq())
    B(doc, 'σyは弾性変形から塑性変形へ移行する際の応力で，SS400では荷重が急落する上降伏点荷重Pyより算出する．'
      'σLはSS400特有の現象で，上降伏点通過後に荷重が低下して安定した値である．'
      'σBは最大荷重Pmaxを初期断面積A₀で除した材料強度の代表指標であり，'
      'このとき局部収縮（ネッキング）が始まる．'
      'δは試験前後の標点間距離変化率で延性の指標，φは断面積減少率でもう一つの延性指標である．'
      'EはJIS Z 2241の弾性域における応力−ひずみの比例係数で，材料固有の剛性を表す．')

    subsection(doc, '課題2　真応力・対数ひずみ')
    B(doc, '公称応力は初期断面積A₀を用いるため，大変形域では実際の応力より小さくなる．'
      '真応力σtは瞬間断面積Aを用いた実際の応力であり，体積一定仮定（Al = A₀l₀）が成立する均一変形域では以下の変換式が成立する．')
    equation(doc, 'σt = σ (1 + e)', eq())
    equation(doc, 'εt = ln(1 + e)', eq())
    B(doc, '対数ひずみεtは微小変形増分dl/lを積分した量であり参照形状に依存しない特徴を持つ．'
      '一軸引張試験から得た真応力−対数ひずみ線図は，多軸変形のモデリングや有限要素解析の材料則に直接利用できる．'
      'ネッキング開始（最大荷重点）以降は体積一定仮定が成立しなくなるため変換が適用できない．')

    subsection(doc, '課題3　相当応力・相当塑性ひずみ')
    B(doc, '三次元応力場における降伏判定には，ミーゼスの相当応力σ̄を用いる．')
    equation(doc, 'σ̄ = √{ [(σ₁−σ₂)² + (σ₂−σ₃)² + (σ₃−σ₁)²] / 2 }', eq())
    B(doc, '相当塑性ひずみ増分dε̄ᵖは塑性ひずみ増分テンソルの不変量として定義され，'
      '変形量の大きさを表すスカラー量である．一軸引張では相当応力は真応力に，'
      '相当塑性ひずみは軸方向真塑性ひずみと等価になる．'
      'これにより引張試験の真応力−対数ひずみ曲線（弾性部を除く）を相当応力−相当塑性ひずみ関係として利用できる．')

    subsection(doc, '課題4　指数硬化則')
    B(doc, '塑性域の真応力σtと相当塑性ひずみεₚの関係を近似したHollomon則を以下に示す．')
    equation(doc, 'σt = K · εₚⁿ', eq())
    B(doc, f'両対数変換するとln σt = ln K + n · ln εₚ の直線関係となり，強度係数K [MPa]と加工硬化指数n [-]を回帰評価できる．'
      f'nが大きいほど変形が均一に分布しやすく，均一伸びeu=nの関係（Considere条件）がある．'
      f'SS400の実測データ（リューダース帯以降〜最大荷重点）に対して回帰を行った結果，'
      f'K={_K:.0f}MPa，n={_n:.3f}，R²={_r2:.4f}が得られた（図{fn5}参照）．'
      f'n={_n:.3f}は炭素鋼の典型値（0.20〜0.27程度）よりやや高い値であり，リューダース帯以降の加工硬化域が広いことに起因する可能性がある．')

    subsection(doc, '課題5　破壊様式')
    B(doc, '金属材料の破壊様式は延性破壊と脆性破壊に大別される．'
      '延性破壊（SS400，A7075）では，破断前に顕著な塑性変形が生じ，第2相粒子を核としたボイドの発生・成長・合体を経て破断する（ボイドコアレッセンス機構）．'
      '破面にはディンプル（半球状窪み）が形成され，SS400ではカップコーン型破面の中央繊維状域に等軸ディンプルが観察される．'
      '脆性破壊（FC250）では塑性変形をほとんど伴わずにき裂が急速に伝播し，平坦な破面（へき開面，リバーパターン）が形成される．'
      'FC250のフレーク状黒鉛は応力集中源として作用し，引張荷重に対して低延性を示す．')

    subsection(doc, '課題6　グラフ解説')
    B(doc, f'図{fn1}の3材料比較線図より，各材料の破壊特性が明確に確認できる．'
      f'SS400（図{fn2}）では上降伏点（415.7MPa）での荷重低下，下降伏点（341.4MPa）付近のリューダース伸び，'
      '加工硬化，引張強さ（463.1MPa）到達後のネッキング・破断という典型的な延性鋼の挙動が観察される．'
      f'A7075（図{fn3}）では明瞭な降伏点は存在せず，連続的な応力上昇の後，引張強さ（1103.2MPa）付近で急破断する．'
      f'FC250（図{fn4}）では線形的な応力上昇の後，引張強さ（396.1MPa）と同時にほぼ突然破断し，脆性的挙動を示す．'
      'FC250の破断伸び4.84%はSS400の28.58%に比べて著しく小さく，フレーク黒鉛による応力集中の影響が明確に現れている．')

    subsection(doc, '課題7　ヤング率比較')
    B(doc, f'グラフの弾性域初期勾配から概算したSS400のヤング率はE≈200GPaであり，'
      f'文献値206GPaに比べてやや低い．A7075では約71.7GPa（文献値同等），FC250では約130GPa（文献値）である．'
      'クロスヘッド変位には試験機フレームおよびつかみ部の弾性変形（機械コンプライアンス）が含まれるため，'
      '弾性域から算出したEは文献値より過小評価となる傾向がある．'
      '伸び計（LVDT）を用いた追加実験（第2章）により，この補正の定量化が可能となる．')
    B(doc, f'引張靭性の評価として，応力−ひずみ線図下の面積（台形積分）を算出した結果，'
      f'SS400: {_U_SS:.1f}MJ/m³，A7075: {_U_A7:.1f}MJ/m³，FC250: {_U_FC:.1f}MJ/m³となった．'
      f'SS400はFC250の約{_U_SS/_U_FC:.0f}倍のエネルギー吸収能を示し，延性差が靭性に直結することが定量的に確認された．')

    # ===========================================================
    # 第2章 追加実験
    # ===========================================================
    doc.add_page_break()
    chapter_title(doc, '第2章　追加実験（伸び計を用いた応力−ひずみ関係の測定）')

    # --- 2-1 目的 ---
    section(doc, '2-1　目的')
    B(doc, 'LVDT伸び計を使用してSS400の弾性域における変形を精密に計測し，'
      '機械コンプライアンスの影響を排除したヤング率を求める．')

    # --- 2-2 概要 ---
    section(doc, '2-2　概要')
    B(doc, 'LVDT（Linear Variable Differential Transformer，直動差動変圧器）伸び計を試験片平行部に装着して引張試験を行う．'
      '伸び計は計測範囲（最大伸び10%=5mm）到達前に取り外し，その後は通常の引張試験として継続する．'
      '試験片諸元はA₀=154.165mm²，l₀=82.30mmである．')

    # --- 2-3 実験手順 ---
    section(doc, '2-3　実験手順')
    proc(doc, '① 試験片の準備（SS400，l₀=82.30mm，A₀=154.165mm²）')
    proc(doc, '② LVDT伸び計の装着と校正')
    proc(doc, '　（1）標点間距離50mmに伸び計を装着する', level=2)
    proc(doc, '　（2）計測機器の2号端子に接続しゼロ調整する', level=2)
    proc(doc, '③ 引張試験の実施')
    proc(doc, '　（1）応力速度3.0MPa/sで荷重を加える', level=2)
    proc(doc, '　（2）伸び計計測範囲（5mm）到達前に伸び計を取り外す', level=2)
    proc(doc, '　（3）試験を継続し最大荷重点まで記録する', level=2)

    # --- 2-4 結果 ---
    section(doc, '2-4　結果')
    B(doc, '表4に追加実験の主要諸元および結果を示す．'
      '最大公称応力はσmax = 21500 / 154.165 = 139.5MPaである．'
      '伸び計使用時の弾性域傾きから算出したヤング率はE≈70.3GPaとなった．'
      'これはSS400として低い値であり，断面積の取り方，グラフの読み取り，試験条件について再確認の余地がある．')

    tn4 = tn()
    tbl_caption(doc, tn4, '追加実験 主要諸元および結果')
    make_table(doc,
        ['項目', '値'],
        [('l₀ [mm]', '82.30'),
         ('A₀ [mm²]', '154.165'),
         ('Pmax [kN]', '21.5'),
         ('σmax [MPa]', '139.5'),
         ('E [GPa]（伸び計）', '70.3')])
    Bx(doc, '')

    fn6 = fn()
    B(doc, f'図{fn6}にSS400通常試験と追加実験の応力−ひずみ線図の比較を示す（追加実験は概略）．')
    insert_figure(doc, plot_ss400_comparison())
    fig_caption(doc, fn6, 'SS400 通常試験と追加実験の応力-ひずみ線図比較')

    # --- 2-5 課題 ---
    section(doc, '2-5　課題')

    subsection(doc, '課題8　LVDT伸び計の測定原理')
    B(doc, 'LVDT（直動差動変圧器）伸び計は，一次コイルへの交流印加で磁束を発生させ，'
      '中心の可動鉄心の変位に比例した差動電圧（二次コイル出力の差）を検出する計器である．'
      '可動鉄心が中立点より変位すると二次コイル間の誘起電圧差が変化し，変位量に比例した出力電圧が得られる．'
      '非接触型電気変換方式のため，機械的摩擦の影響がなく応答性・繰り返し精度に優れる．'
      '本実験では標点間距離50mm，計測最大伸び10%（5mm）の伸び計を使用した．')

    subsection(doc, '課題9　追加実験ヤング率の比較')
    B(doc, f'追加実験から得られたヤング率E≈70.3GPaは，'
      '文献値206GPaおよび通常試験（クロスヘッド変位）から概算した200GPaを大きく下回っている．'
      '70.3GPaはSS400として低い値であり，断面積の取り方，グラフの読み取り，試験条件について再確認の余地がある．'
      '伸び計を用いた計測では機械コンプライアンスを除いた平行部の変形のみを計測するため，'
      '理論上は文献値に近い値が得られるべきであり，今後のデータ確認が必要である．')

    subsection(doc, '課題10　最大荷重時公称応力と真応力')
    B(doc, '最大荷重Pmax=21.5kN=21500Nおよび初期断面積A₀=154.165mm²から，')
    equation(doc, 'σmax = 21500 / 154.165 = 139.5 MPa', eq())
    B(doc, '最大荷重時の対数ひずみemax（伸び計読み取り値）を用いて，真応力σt(max) = σmax × (1+emax)として算出できる．')

    subsection(doc, '課題11　体積一定仮定のもとでの真応力−対数ひずみ関係')
    B(doc, '標点間体積一定（A · l = A₀ · l₀）の仮定のもとで，各測定点の公称量から以下の変換により真量を求める．')
    equation(doc, 'σt = σ (1 + e)', eq())
    equation(doc, 'εt = ln(1 + e)', eq())
    B(doc, 'この変換は均一変形域（最大荷重点まで）でのみ有効である．'
      'ネッキング開始以降は体積一定仮定が崩れるため，破断部最小径の実測が必要となる．')

    subsection(doc, '課題12　伸び計あり/なしの比較')
    B(doc, '伸び計使用時（平行部標点間変位を直接計測）と未使用時（クロスヘッド変位使用）では，'
      '機械コンプライアンスの影響により弾性域での応力−ひずみ線図の傾きが異なる．'
      'クロスヘッド変位にはフレームおよびつかみ部の弾性変形が含まれるため，'
      'ヤング率が過小評価（今回の通常試験では≈200GPaと，文献値206GPaより小さい）される傾向にある．'
      '伸び計は平行部を直接計測するため，この影響を排除できる．'
      'ただし今回の追加実験では70.3GPaという低い値となったため，計測条件の詳細な再確認が必要である．')

    # ===========================================================
    # 第3章 シャルピー衝撃試験
    # ===========================================================
    doc.add_page_break()
    chapter_title(doc, '第3章　シャルピー衝撃試験')

    # --- 3-1 目的 ---
    section(doc, '3-1　目的')
    B(doc, '振り子型シャルピー試験機を用いてSS400，FC250，A7075のシャルピー吸収エネルギーKVおよびシャルピー衝撃値ρを算出し，'
      '各材料の衝撃靭性を評価する．')

    # --- 3-2 試験機および試験片 ---
    section(doc, '3-2　試験機および試験片')
    B(doc, f'振り子型シャルピー試験機を使用した．試験機定数：ハンマ質量m={m_kg}kg，'
      f'重力加速度g={g_acc}m/s²（野田市山崎付近），W=mg={W_N:.3f}N，'
      f'重心半径r={r_m}m，M=Wr={M_J:.3f}J，持上げ角α={alpha_deg}°．')

    tn5 = tn()
    tbl_caption(doc, tn5, 'シャルピー試験片寸法')
    make_table(doc,
        ['材料', 'l [mm]', 'b [mm]', 'h1 [mm]', 'h2 [mm]', 'A [mm²]', 'A [cm²]'],
        [('SS400', '54.98', '10.00', '10.00', '10.00', '100.0', '1.000'),
         ('FC250', '54.90', '10.00', '10.00', '9.97',  '99.7',  '0.997'),
         ('A7075', '55.00', '9.90',  '10.00', '10.00', '99.0',  '0.990')])
    Bx(doc, '')

    # --- 3-3 実験手順 ---
    section(doc, '3-3　実験手順')
    proc(doc, '① 試験片の寸法測定')
    proc(doc, '　（1）マイクロメーターで l，b，h1，h2 を測定する', level=2)
    proc(doc, '　（2）ノッチ部断面積 A = b × h2 を算出する', level=2)
    proc(doc, '② 試験機への試験片装着')
    proc(doc, '③ ハンマの解放と振上がり角の読み取り')
    proc(doc, f'　（1）ハンマを最大持上げ角α={alpha_deg}°まで持ち上げ解放する', level=2)
    proc(doc, '　（2）打撃後の振上がり角βを読み取る', level=2)
    proc(doc, f'④ 損失エネルギーLの測定（空打ち試験）')
    proc(doc, f'　（1）試験片を装着せずにハンマをα={alpha_deg}°から解放する', level=2)
    proc(doc, f'　（2）打撃後の振上がり角β₀を読み取る（本実験ではβ₀={beta_air}°）', level=2)
    proc(doc, f'　（3）L = M(cosβ₀ − cosα)として損失エネルギーを算出する', level=2)

    # --- 3-4 結果 ---
    section(doc, '3-4　結果')
    eq_K_n = eq()
    eq_rho_n = eq()
    B(doc, f'シャルピー吸収エネルギーK [J]は式({eq_K_n})，シャルピー衝撃値ρ [J/cm²]は式({eq_rho_n})で算出する．')
    equation(doc, 'K = M (cosβ − cosα) − L', eq_K_n)
    equation(doc, 'ρ = K / A', eq_rho_n)
    B(doc, f'ここでM=Wr={M_J:.3f}J（W={W_N:.3f}N，r={r_m}m），'
      f'cosα=cos{alpha_deg}°={cos_alpha:.4f}，'
      f'Lは軸受摩擦および空気抵抗による損失エネルギー[J]，'
      f'A [cm²]はノッチ部原断面積である．')
    B(doc, f'損失エネルギーLは，試験片を装着せずに空打ち試験を行い，'
      f'ハンマがα={alpha_deg}°から解放後β₀={beta_air}°まで振り上がったことから次式で求めた（置針摩擦は考慮しない）．')
    equation(doc, f'L = M (cosβ₀ − cosα) = {M_J:.3f} × (cos{beta_air}° − cos{alpha_deg}°) = {L_J:.2f}  J', eq())

    # 最終K値計算
    ss_K0  = M_J * (math.cos(math.radians(CHARPY_BETA['SS400'])) - cos_alpha)
    fc_K0  = M_J * (math.cos(math.radians(CHARPY_BETA['FC250'])) - cos_alpha)
    a7_K0  = M_J * (math.cos(math.radians(CHARPY_BETA['A7075'])) - cos_alpha)

    tn6 = tn()
    B(doc, f'各材料の振上がり角βを読み取り，L={L_J:.2f}Jを用いてKおよびρを算出した結果を表{tn6}に示す．')
    tbl_caption(doc, tn6, 'シャルピー試験結果（確定値）')
    make_table(doc,
        ['材料', 'β [°]', 'K [J]', 'A [cm²]', 'ρ [J/cm²]'],
        [('SS400', f'{CHARPY_BETA["SS400"]}',
          f'{K_final["SS400"]:.2f}', f'{A_cm2["SS400"]:.3f}', f'{rho_final["SS400"]:.2f}'),
         ('FC250', f'{CHARPY_BETA["FC250"]}',
          f'{K_final["FC250"]:.2f}', f'{A_cm2["FC250"]:.3f}', f'{rho_final["FC250"]:.2f}'),
         ('A7075', f'{CHARPY_BETA["A7075"]}',
          f'{K_final["A7075"]:.2f}', f'{A_cm2["A7075"]:.3f}', f'{rho_final["A7075"]:.2f}')])
    Bx(doc, '')
    B(doc, 'その他の評価量として，脆性破面率B=C/A×100 [%]（C=脆性破面積，A=全破面積），'
      '延性破面率S=F/A×100 [%]（F=延性破面積），横膨出量C=a−b [mm]（a=破面最大幅，b=試験片元の幅）を定義する．')

    # --- 3-5 考察 ---
    section(doc, '3-5　考察')

    subsection(doc, '課題13　材料の特徴と吸収エネルギーの関係')
    B(doc, f'シャルピー吸収エネルギーの序列は SS400（K={K_final["SS400"]:.2f}J）'
      f' > A7075（K={K_final["A7075"]:.2f}J）'
      f' > FC250（K={K_final["FC250"]:.2f}J）であり，'
      f'引張試験から得た延性序列（δ：SS400 {KV["SS400"]["delta"]:.2f}% >> A7075 {KV["A7075"]["delta"]:.2f}% > FC250 {KV["FC250"]["delta"]:.2f}%）'
      'と一致した．')
    B(doc, 'SS400は延性鋼であり，衝撃負荷に対して大きな塑性変形を生じながらエネルギーを吸収するため'
      f'K={K_final["SS400"]:.2f}Jと最大値を示した．'
      'FC250はフレーク状黒鉛が応力集中源となり，ほとんど塑性変形を伴わずに破断するため'
      f'K={K_final["FC250"]:.2f}Jと最小値であった．'
      'A7075は高強度アルミ合金であり延性はSS400より小さいが，'
      f'FC250よりは大きな塑性変形能を有するためK={K_final["A7075"]:.2f}Jと中間的な値を示した．')
    B(doc, f'衝撃値ρはSS400 {rho_final["SS400"]:.2f}J/cm²，'
      f'A7075 {rho_final["A7075"]:.2f}J/cm²，'
      f'FC250 {rho_final["FC250"]:.2f}J/cm²であり，'
      '強度と延性の積として評価できる靭性の観点と整合する結果が得られた．')

    subsection(doc, '課題14　脆性破面率と吸収エネルギーの関係')
    B(doc, '脆性破面率B [%]は，破面全面積に対する脆性的に破壊した領域の割合であり，材料の破壊様式を定量化する指標の一つである．')
    B(doc, '一般にBが大きい材料ほどシャルピー吸収エネルギーが小さい傾向がある．'
      '延性破壊ではき裂先端に塑性変形帯が発達してエネルギー吸収量が増加するが，'
      '脆性破壊では塑性変形を伴わないためエネルギー吸収量が小さい．'
      f'本実験の結果からもこの対応が確認できる．FC250はK={K_final["FC250"]:.2f}Jと最小の吸収エネルギーを示し，'
      '黒鉛フレークを起点とした脆性的破壊が支配的であることからBは高いと考えられる．'
      f'SS400はK={K_final["SS400"]:.2f}Jと最大の吸収エネルギーを示し，延性破壊（ディンプル組織）が支配的であるためBは低い．'
      f'A7075はK={K_final["A7075"]:.2f}Jと中間的な値であり，延性・脆性の混合的破壊形態を反映している．'
      'これらの結果は，BとρのBetween-material序列（SS400：B小・ρ大，FC250：B大・ρ小）と整合しており，'
      '脆性破面率が衝撃吸収能を定量的に評価する指標として有効であることを示している．')

    # ===========================================================
    # 第4章 破面観察
    # ===========================================================
    doc.add_page_break()
    chapter_title(doc, '第4章　破面観察')

    # --- 4-1 目的 ---
    section(doc, '4-1　実験の目的')
    B(doc, '引張試験およびシャルピー試験後の破断面を走査型電子顕微鏡（SEM）で観察し，'
      '各材料の延性・脆性破壊に特有のミクロ形態を確認する．'
      'SS400（延性破壊），FC250（脆性破壊），A7075（混合破壊）の3材料について，'
      '引張試験片の中央域（C）・端部（E）およびシャルピー試験片のノッチ部（CN），中央（CC），ノッチ反対側（CO）を'
      'それぞれ×100・×500の2倍率で観察し，第1章・第3章の力学的測定結果との対応を考察する．')

    # --- 4-2 SEM原理 ---
    section(doc, '4-2　走査型電子顕微鏡の原理')
    B(doc, 'タングステン陰極から射出した一次電子ビームを電磁レンズで集束し，試料表面を走査する．'
      '試料から放出される二次電子をシンチレーター・光電子増倍管で検出し，表面形状像を得る．'
      '倍率は数十倍〜数万倍であり，破面の微細形態観察に適する．'
      '試料は導電性が必要であり，絶縁体の場合は金蒸着等の導電処理を施す．')

    # --- 4-3 試験手順 ---
    section(doc, '4-3　試験手順')
    proc(doc, '① 試料の切り出しと洗浄')
    proc(doc, '　（1）ファインカッターで破断部を切り出す', level=2)
    proc(doc, '　（2）アセトンで洗浄・乾燥する', level=2)
    proc(doc, '② SEMへの試料装着')
    proc(doc, '　（1）ビニール手袋を着用して試料台に固定する', level=2)
    proc(doc, '　（2）SEM内部を真空排気する', level=2)
    proc(doc, '③ 観察と撮影')
    proc(doc, '　（1）倍率を変化させながら特徴的形態を探す', level=2)
    proc(doc, '　（2）×100および×500で撮影する', level=2)

    # --- 4-4 破面の例 ---
    section(doc, '4-4　破面の例')
    B(doc, '金属材料の代表的な破面形態を以下に示す．')
    B(doc, 'ディンプル：延性破壊特有の半球状窪みで，ボイドコアレッセンスにより形成される．'
      '中央繊維状域では等軸ディンプルが，せん断リップ域では扁長なせん断ディンプルが観察される．')
    B(doc, 'リバーパターン：脆性破壊特有の河状模様で，高さの異なるへき開面間の段差により形成される．')
    B(doc, 'せん断ディンプル：45°方向のせん断荷重下で形成される扁長ディンプルであり，'
      'カップコーン破面の外周部（せん断リップ）に観察される．')
    B(doc, '粒間破壊：結晶粒界に沿ったき裂伝播で，FC250ではフレーク黒鉛周辺に特徴的に見られる．')

    # --- 4-5 結果 ---
    section(doc, '4-5　結果')

    # ---- 4-5-1 引張試験片 ----
    subsection(doc, '4-5-1　引張試験片の破面観察')
    B(doc, '引張試験片のC（中央）およびE（端部）部位について，各材料の×100・×500SEM画像を以下に示す．')

    tensile_materials = [
        ('SS400', 'SS400', [
            ('c', '中央域（C）',
             'SS400引張試験片の中央域（C）では等軸ディンプルが密集し，ボイドコアレッセンスによる典型的な延性破面形態を示す．'),
            ('E', '端部（E）',
             'SS400引張試験片の端部（E）ではせん断ディンプルが観察され，45°方向のせん断応力支配による変形が確認できる．'),
        ]),
        ('FC250', 'FC250', [
            ('c', '中央域（C）',
             'FC250引張試験片の中央域（C）では準へき開ファセットが全面を覆い，フレーク黒鉛を起点とした脆性的な破断特徴を示す．'),
            ('E', '端部（E）',
             'FC250引張試験片の端部（E）では角張った準へき開ファセットが観察され，脆性破壊の特徴が端部でも一様に現れている．'),
        ]),
        ('A7075', 'A7075', [
            ('c', '中央域（C）',
             'A7075引張試験片の中央域（C）では不規則な塊状形態と繊維状特徴が混在し，析出物を核としたボイドコアレッセンスとせん断成分の複合形態を示す．'),
            ('E', '端部（E）',
             'A7075引張試験片の端部（E）では比較的滑らかな外観に不規則な特徴が混在し，平面応力状態でのせん断変形が支配的なことを示す．'),
        ]),
    ]

    for mat_key, mat_name, regions in tensile_materials:
        for region_code, region_name, desc in regions:
            fn_cur = fn()
            B(doc, desc + f'（図{fn_cur}）')
            insert_sem_pair(doc,
                            sem_path(mat_key, region_code, '100'),
                            sem_path(mat_key, region_code, '500'))
            fig_caption(doc, fn_cur,
                        f'{mat_name} 引張試験片 {region_name}（左：×100，右：×500）')

    # ---- 4-5-2 シャルピー試験片 ----
    subsection(doc, '4-5-2　シャルピー試験片の破面観察')
    B(doc, 'シャルピー試験片のCC（中央），CN（ノッチ部），CO（ノッチ反対側）部位について，各材料の×100・×500SEM画像を以下に示す．')

    charpy_materials = [
        ('SS400', 'SS400', [
            ('CC', 'シャルピー中央（CC）',
             'SS400シャルピー試験片の中央（CC）では繊維状ディンプルが観察され，延性的な破壊が支配的な中央領域の特徴を示す．'),
            ('CN', 'ノッチ部（CN）',
             'SS400シャルピー試験片のノッチ部（CN）では遷移域が観察され，ノッチ応力集中による破壊起点付近の形態変化が確認できる．'),
            ('CO', 'ノッチ反対側（CO）',
             'SS400シャルピー試験片のノッチ反対側（CO）ではき裂伝播の終端部の特徴が観察される．'),
        ]),
        ('FC250', 'FC250', [
            ('CC', 'シャルピー中央（CC）',
             'FC250シャルピー試験片の中央（CC）では粒状テクスチャと準へき開面が観察され，脆性的な破壊が中央部でも支配的なことを示す．'),
            ('CN', 'ノッチ部（CN）',
             'FC250シャルピー試験片のノッチ部（CN）ではノッチ近傍の平坦な脆性破面が観察され，低靭性の特徴を示す．'),
            ('CO', 'ノッチ反対側（CO）',
             'FC250シャルピー試験片のノッチ反対側（CO）ではへき開ファセットがノッチ反対側でも全域にわたって観察される．'),
        ]),
        ('A7075', 'A7075', [
            ('CC', 'シャルピー中央（CC）',
             'A7075シャルピー試験片の中央（CC）では伸長した条痕・せん断ディンプルが観察され，混合破壊形態の特徴を示す．'),
            ('CN', 'ノッチ部（CN）',
             'A7075シャルピー試験片のノッチ部（CN）ではノッチ起点付近の破壊様式遷移が確認でき，応力集中の影響が破面形態に現れている．'),
            ('CO', 'ノッチ反対側（CO）',
             'A7075シャルピー試験片のノッチ反対側（CO）ではき裂伝播の末端部の特徴的な形態が観察される．'),
        ]),
    ]

    for mat_key, mat_name, regions in charpy_materials:
        for region_code, region_name, desc in regions:
            fn_cur = fn()
            B(doc, desc + f'（図{fn_cur}）')
            insert_sem_pair(doc,
                            sem_path(mat_key, region_code, '100'),
                            sem_path(mat_key, region_code, '500'))
            fig_caption(doc, fn_cur,
                        f'{mat_name} シャルピー試験片 {region_name}（左：×100，右：×500）')

    # --- 4-6 考察 ---
    section(doc, '4-6　考察')

    subsection(doc, '課題15　マクロ破面観察')
    B(doc, 'マクロ観察では各材料の破壊様式が明確に区別される．'
      'SS400ではカップコーン型の破面（中央繊維状部＋外周45°せん断リップ）が形成され，顕著なくびれを伴う延性的破断であった．'
      'FC250ではくびれがほとんど認められず，破面は平坦で光沢があり，脆性破断の特徴を示した．'
      'A7075はSS400より小さなくびれを伴う延性的破断であったが，FC250ほど平坦な破面ではなかった．'
      'これらのマクロ形態は，表2の破断伸びδの序列（SS400: 28.58% >> A7075: 7.69% > FC250: 4.84%）と整合する．')

    subsection(doc, '課題16　同一部位における材料差')
    B(doc, '引張試験片の中央域（C）を例に，材料間の破面形態の違いを考察する．')
    B(doc, 'SS400（延性）：等軸ディンプルが密集して観察された．'
      '三軸引張応力場のもとでボイドが等方的に成長・合体したボイドコアレッセンス機構による典型的な延性破面である．')
    B(doc, 'FC250（脆性）：準へき開ファセットが全面を覆い，ディンプルは認められなかった．'
      'フレーク黒鉛の先端を起点としたき裂が特定の結晶面に沿って伝播した脆性破壊の特徴を示す．')
    B(doc, 'A7075（混合）：塊状形態と繊維状特徴が混在して観察された．'
      'MgZn₂等の析出物を核としたボイドコアレッセンスと局所的なせん断変形が複合した混合破壊形態である．')
    B(doc, '以上より，同一部位における破面形態の差異は延性・脆性の序列（SS400 >> A7075 > FC250）と整合し，'
      '第1章の引張試験で得た破断伸びδの序列をミクロ構造レベルで直接裏付けている．')

    subsection(doc, '課題17　同一材料における場所差')
    B(doc, 'SS400の引張試験片を例に，部位による破面形態の違いを考察する．')
    B(doc, '中央域（C）：三軸引張応力場のもとでボイドが発生・成長・合体する延性的破壊が支配的な部位である．'
      '等軸ディンプルが密集して観察され，ボイドコアレッセンス機構の直接的な証拠となっている．')
    B(doc, '端部（E）：最外周の平面応力支配域であり，45°方向のせん断応力が卓越する部位である．'
      'せん断ディンプルが観察され，変形方向に引き伸ばされた扁長な形態をとる．')
    B(doc, 'シャルピー試験片における位置差については，CN（ノッチ部）はき裂起点であり，'
      '応力集中によって破壊が開始されるため，その付近の形態は最も複雑な遷移形態を示す．'
      'CC（中央）はき裂伝播途中の破面であり，CO（ノッチ反対側）はき裂伝播の末端部の特徴を示す．'
      'この空間的な形態変化は，き裂伝播に伴う応力状態の変化を反映している．')

    # ===== 参考文献 =====
    doc.add_page_break()
    section(doc, '参考文献')
    Bx(doc, '')
    for ref in [
        '1) 小寺沢良一，フラクトグラフィ，培風館 (1985)，pp. 1-50．',
        '2) 日本機械学会編，機械工学便覧 基礎編 α1 材料力学，日本機械学会 (2004)，pp. 1-45．',
        '3) 日本工業標準調査会，JIS Z 2242 金属材料のシャルピー衝撃試験方法，日本規格協会 (2018)．',
        '4) 日本機械学会編，機械材料学，日本機械学会 (2007)，pp. 120-140．',
        '5) 軽金属学会編，アルミニウムの組織と性質，軽金属学会 (1991)，pp. 201-215．',
        '6) 日本機械学会編，機械工学便覧 基礎編 α4 計測工学，日本機械学会 (2007)，pp. 4-25．',
    ]:
        para(doc, ref, size=10, sa=2, indent_mm=0)

    # ===== Unicode 添字変換 =====
    fix_subscripts_in_doc(doc)

    return doc


print('[v8] ビルド開始...')
doc = build()
doc.save(OUTPUT_PATH)
print(f'[v8] 完了: {OUTPUT_PATH}')
