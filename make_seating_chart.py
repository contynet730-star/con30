import sys, copy
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
ws = wb['座席表CS入力済']

# group definitions: order follows グループ一覧 (08), names as in 地下多目 (10)
G = {
 '幼稚園':        ('幼', ['北大泉','光が丘むらさき','光が丘さくら']),
 '小 豊玉地区':   ('小', ['みらい青空学園（小）','小竹小','豊玉小','豊玉第二小','豊玉東小','豊玉南小','中村小','中村西小']),
 '小 開進地区':   ('小', ['早宮小','開進第一小','開進第二小','開進第三小','開進第四小','仲町小','南町小','北町小','北町西小']),
 '小 練馬地区':   ('小', ['練馬小','練馬第二小','練馬第三小','練馬東小','田柄小','田柄第二小','向山小','春日小']),
 '小 光が丘地区': ('小', ['旭町小','高松小','光が丘四季小','光が丘春風小','光が丘夏雲小','光が丘秋陽小','光が丘第八小']),
 '小 石神井東地区':('小', ['石神井東小','下石神井小','光和小','谷原小','北原小','南田中小','南が丘小','富士見台小']),
 '小 石神井西地区':('小', ['石神井小','石神井西小','石神井台小','上石神井小','上石北小','立野小','関町小','関町北小']),
 '小 大泉東地区': ('小', ['豊溪小','大泉第一小','大泉東小','大泉北小','大泉学園小','泉新小','橋戸小','八坂小']),
 '小 大泉西地区': ('小', ['大泉小','大泉第二小','大泉第三小','大泉第四小','大泉第六小','大泉西小','大泉南小','大泉学園緑小','大泉桜学園（小）']),
 '中 Ａ地区':     ('中', ['みらい青空学園（中）','豊玉中','豊玉第二中','中村中','開進第一中','開進第二中','開進第三中','開進第四中','北町中']),
 '中 Ｂ地区':     ('中', ['練馬中','練馬東中','貫井中','田柄中','豊渓中','光が丘第一中','光が丘第二中','光が丘第三中']),
 '中 Ｃ地区':     ('中', ['谷原中','三原台中','大泉中','大泉西中','大泉北中','大泉学園中','大泉桜学園（中）','八坂中']),
 '中 Ｄ地区':     ('中', ['石神井中','石神井東中','石神井西中','石神井南中','上石神井中','南が丘中','大泉第二中','関中']),
}
assert sum(len(v[1]) for v in G.values()) == 101

COLS = ['C','D','E','G','H','I','K','L','M','O','P','Q']   # seat columns ①..⑫
ROWS = [5,7,10,12,15,17,20,22,25]                           # anchor rows of seat rows ①..⑨
FILL = {'幼': 'CCECFF', '小': 'CCFFCC', '中': 'FFCCE5'}
red = Side(style='thick', color='FF0000')
from openpyxl.utils import column_index_from_string as ci

# placement: group -> halves, each half = list of (row, col_from, col_to), 1-based; filled left->right, top->bottom
PLACE = [
 ('幼稚園',         [[(1, 1, 3)]]),
 ('小 開進地区',    [[(1, 4, 8)], [(1, 9, 12)]]),
 ('小 豊玉地区',    [[(2, 1, 4)], [(2, 5, 8)]]),
 ('小 練馬地区',    [[(2, 9, 12)], [(3, 9, 12)]]),
 ('小 石神井東地区',[[(3, 1, 4)], [(3, 5, 8)]]),
 ('小 石神井西地区',[[(4, 1, 4)], [(4, 5, 8)]]),
 ('小 光が丘地区',  [[(4, 9, 12)], [(5, 10, 12)]]),
 ('小 大泉西地区',  [[(5, 1, 5)], [(5, 6, 9)]]),
 ('小 大泉東地区',  [[(6, 1, 4)], [(6, 5, 8)]]),
 ('中 Ｂ地区',      [[(6, 9, 12)], [(7, 9, 12)]]),
 ('中 Ｃ地区',      [[(7, 1, 4)], [(7, 5, 8)]]),
 ('中 Ａ地区',      [[(8, 1, 4)], [(9, 1, 5)]]),
 ('中 Ｄ地区',      [[(8, 5, 8)], [(8, 9, 12)]]),
]
owner, seatname, listing = {}, {}, []
for g, halves in PLACE:
    names = G[g][1]; k = 0
    for hi, spans in enumerate(halves):
        seats = sorted((r-1, c-1) for r, c1, c2 in spans for c in range(c1, c2+1))
        for (r, c) in seats:
            assert (r, c) not in owner
            owner[(r, c)] = (g, hi)
            seatname[(r, c)] = names[k]
            listing.append((g, hi, k+1, names[k], "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"[c] + '列', "①②③④⑤⑥⑦⑧⑨"[r] + '番'))
            k += 1
    assert k == len(names), g
assert len(listing) == 101
# empty seats only in the back row, front row full
assert all((r, c) in owner for r in range(8) for c in range(12))

SHORT = {'幼稚園':'幼稚園','小 豊玉地区':'2 豊玉地区','小 開進地区':'3 開進地区','小 練馬地区':'4 練馬地区',
         '小 光が丘地区':'5 光が丘地区','小 石神井東地区':'6 石神井東地区','小 石神井西地区':'7 石神井西地区',
         '小 大泉東地区':'8 大泉東地区','小 大泉西地区':'9 大泉西地区','中 Ａ地区':'10 中Ａ地区',
         '中 Ｂ地区':'11 中Ｂ地区','中 Ｃ地区':'12 中Ｃ地区','中 Ｄ地区':'13 中Ｄ地区'}
split = Side(style='mediumDashed', color='FF0000')     # line between the two halves of one group
def setb(cell, **kw):
    o = cell.border
    d = dict(left=o.left, right=o.right, top=o.top, bottom=o.bottom); d.update(kw)
    cell.border = Border(**d)

def draw(ws, mode, title, legend):
    """mode 'normal': one red box per group / 'half': each group split into two halves (dashed line)."""
    unit = (lambda o: o) if mode == 'half' else (lambda o: o and o[0])
    own = lambda r, c: unit(owner.get((r, c))) if 0 <= r < 9 and 0 <= c < 12 else None
    def side(a, b):
        if b == a: return None
        if mode == 'half' and b is not None and b[0] == a[0]: return split
        return red
    def grp(u): return u if isinstance(u, str) else u[0]
    for r in range(9):
        for c in range(12):
            col = COLS[c]; row = ROWS[r]; u = own(r, c); name = seatname.get((r, c))
            top = ws[f'{col}{row}']
            if name:
                o = owner[(r, c)]
                label = SHORT[o[0]] + ('' if o[0] == '幼稚園' or mode == 'normal' else '①②'[o[1]] + '班')
                n = len(name); fs = 10 if n <= 7 else 9 if n == 8 else 7
                top.value = CellRichText(
                    TextBlock(InlineFont(rFont='ＭＳ ゴシック', sz=fs), name + '\n'),
                    TextBlock(InlineFont(rFont='ＭＳ ゴシック', sz=12, color='FFFF0000', b=True), label))
            else:
                top.value = None
            top.font = Font(name='ＭＳ ゴシック', size=11)
            top.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            for cell in (top, ws[f'{col}{row+1}']):
                cell.fill = PatternFill('solid', fgColor=FILL[G[grp(u)][0]]) if u else PatternFill(fill_type=None)
            if u is None: continue
            sl, sr = side(u, own(r, c-1)), side(u, own(r, c+1))
            st, sb = side(u, own(r-1, c)), side(u, own(r+1, c))
            ccol = ci(col)
            for k, rr in enumerate((row, row+1)):
                kw = {}
                if sl: kw['left'] = sl; setb(ws.cell(rr, ccol-1), right=sl)
                if sr: kw['right'] = sr; setb(ws.cell(rr, ccol+1), left=sr)
                if st and k == 0: kw['top'] = st
                if sb: kw['bottom'] = sb
                setb(ws.cell(rr, ccol), **kw)
            if st: setb(ws.cell(row-1, ccol), bottom=st)
            if sb: setb(ws.cell(row+2, ccol), top=sb)
            gfill = PatternFill('solid', fgColor=FILL[G[grp(u)][0]])
            # unit continues across an aisle column to the right
            if c < 11 and own(r, c+1) == u and ci(COLS[c+1]) - ccol == 2:
                acol = ccol + 1
                for rr in (row, row+1): ws.cell(rr, acol).fill = gfill
                if st: setb(ws.cell(row, acol), top=st); setb(ws.cell(row-1, acol), bottom=st)
                if sb: setb(ws.cell(row+1, acol), bottom=sb); setb(ws.cell(row+2, acol), top=sb)
            # unit continues across an aisle row below
            if r < 8 and own(r+1, c) == u and ROWS[r+1] - row == 3:
                arow = row + 2
                ws.cell(arow, ccol).fill = gfill
                if c < 11 and ci(COLS[c+1]) - ccol == 2 and own(r, c+1) == u and own(r+1, c+1) == u:
                    ws.cell(arow, ccol+1).fill = gfill
                lb = side(u, own(r, c-1)) or side(u, own(r+1, c-1))
                rb = side(u, own(r, c+1)) or side(u, own(r+1, c+1))
                if lb: setb(ws.cell(arow, ccol), left=lb); setb(ws.cell(arow, ccol-1), right=lb)
                if rb: setb(ws.cell(arow, ccol), right=rb); setb(ws.cell(arow, ccol+1), left=rb)
    for c in 'CDEGHIKLMOPQ':
        ws.column_dimensions[c].width = 19.0
    ws['B1'].value = title
    ws['B28'].value = legend
    ws['B28'].font = Font(name='ＭＳ ゴシック', size=10, color='FF0000')
    ws['B28'].alignment = Alignment(horizontal='left', vertical='center')
    ws.print_area = 'A1:R29'
    ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

base_title = ws['B1'].value.strip()
ws_half = wb.copy_worksheet(ws)
ws.title = '座席表（普通）'
ws_half.title = '座席表（半分）'
wb.move_sheet(ws_half, offset=-(len(wb.sheetnames) - 2))
draw(ws, 'normal', base_title + '　グループ別',
     '赤枠＝グループ　青：幼稚園　緑：小学校　桃：中学校　白：空席（最後尾）')
draw(ws_half, 'half', base_title + '　グループ別（半分）',
     '太い赤枠＝グループ／赤点線＝グループ内の２班（①班・②班）　青：幼稚園　緑：小学校　桃：中学校　白：空席（最後尾）')

# group summary sheet
NUM = {'幼稚園':'１ 幼稚園','小 豊玉地区':'２ 小学校 豊玉地区','小 開進地区':'３ 小学校 開進地区','小 練馬地区':'４ 小学校 練馬地区',
       '小 光が丘地区':'５ 小学校 光が丘地区','小 石神井東地区':'６ 小学校 石神井東地区','小 石神井西地区':'７ 小学校 石神井西地区',
       '小 大泉東地区':'８ 小学校 大泉東地区','小 大泉西地区':'９ 小学校 大泉西地区','中 Ａ地区':'10 中学校 Ａ地区',
       '中 Ｂ地区':'11 中学校 Ｂ地区','中 Ｃ地区':'12 中学校 Ｃ地区','中 Ｄ地区':'13 中学校 Ｄ地区'}
order = list(G.keys())
listing.sort(key=lambda t: (order.index(t[0]), t[2]))
ls = wb.create_sheet('グループ別一覧', 2)
ls['A1'] = '11月26日（木）座席表　グループ別一覧'
ls['A1'].font = Font(name='ＭＳ ゴシック', size=12, bold=True)
ls.append([])
ls.append(['グループ', '班', '№', '学校・園名', '座席（列）', '座席（番）'])
H = 3
thin = Side(style='thin')
for c in ls[H]:
    c.font = Font(name='ＭＳ ゴシック', size=11, bold=True); c.fill = PatternFill('solid', fgColor='DDDDDD')
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin); c.alignment = Alignment(horizontal='center', vertical='center')
r = H + 1
for gi, g in enumerate(order):
    rows = [t for t in listing if t[0] == g]
    start = r
    for t in rows:
        vals = [NUM[g], ('－' if g == '幼稚園' else '①②'[t[1]]), t[2], t[3], t[4], t[5]]
        for j, v in enumerate(vals, 1):
            c = ls.cell(r, j, v if (j > 1 or r == start) else None)
            c.font = Font(name='ＭＳ ゴシック', size=11)
            c.fill = PatternFill('solid', fgColor=FILL[G[g][0]])
            c.alignment = Alignment(horizontal='left' if j == 4 else 'center', vertical='center', wrap_text=(j == 1))
            c.border = Border(left=red if j == 1 else thin, right=red if j == 6 else thin,
                              top=red if r == start else (split if t[1] == 1 and rows[r-start-1][1] == 0 else thin),
                              bottom=red if r == start + len(rows) - 1 else thin)
        r += 1
    ls.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
for col, w in zip('ABCDEF', [24, 5, 5, 24, 11, 11]): ls.column_dimensions[col].width = w
ls.freeze_panes = 'A4'
ls.page_setup.orientation = 'portrait'; ls.page_setup.paperSize = 9
ls.page_setup.fitToWidth = 1; ls.page_setup.fitToHeight = 0
ls.sheet_properties.pageSetUpPr.fitToPage = True
ls.print_title_rows = '3:3'
wb.active = 0
wb.save(out)
print(len(listing), 'seats filled')
