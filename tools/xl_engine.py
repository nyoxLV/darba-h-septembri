# -*- coding: utf-8 -*-
"""Bez-Excel dzinējs "darba stundas" .xlsm failam (INSTRUKCIJA 8. nod. alternatīva).

Strādā tieši ar XML (xl/worksheets/sheet1.xml u.c.), nevis caur openpyxl, tāpēc:
  * visas pārējās šūnas, piezīmes, VBA, definētie nosaukumi paliek baitu līmenī neaizskarti;
  * formulu rezultāti (<v>) tiek pārrēķināti Pythonā un ierakstīti atpakaļ, tāpēc fails
    atveras ar pareizām summām arī tad, ja Excel makro ir bloķēti.

Atbalstītās formulas (vienīgās, kas ir lapā "Darbs - <Mēnesis>"):
  SumByHeaderColor(hdr, range)  - imitē Module1: saskaita skaitļus, kuru fona krāsa == galvenes krāsa
  SUM, SUMIF(range, "teksts", sum_range), + - * / ^ & salīdzinājumi, iekavas.
Pārbaude: `verify()` pārrēķina VISAS formulas un salīdzina ar Excel saglabātajām vērtībām.
"""
from __future__ import annotations

import html
import math
import re

CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def col2n(s: str) -> int:
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class XLError(Exception):
    pass


def fmt_num(x: float) -> str:
    if isinstance(x, bool):
        return "1" if x else "0"
    if float(x).is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(float(x))


# ----------------------------------------------------------------------------- krāsas
DEFAULT_INDEXED = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    "000000", "FFFFFF",  # 64 = system foreground, 65 = system background
]


# Excel tēmas krāsu tonējums: Windows ColorRGBToHLS/ColorHLSToRGB (HLSMAX=240, veseli skaitļi).
# Pārbaudīts pret Excel saglabātajām SumByHeaderColor vērtībām (piem. accent6 +40% = A9D08E).
HLSMAX, RGBMAX = 240, 255


def _rgb2hls(R, G, B):
    cmax, cmin = max(R, G, B), min(R, G, B)
    L = (((cmax + cmin) * HLSMAX) + RGBMAX) // (2 * RGBMAX)
    if cmax == cmin:
        return HLSMAX * 2 // 3, L, 0
    if L <= HLSMAX // 2:
        S = (((cmax - cmin) * HLSMAX) + ((cmax + cmin) // 2)) // (cmax + cmin)
    else:
        S = (((cmax - cmin) * HLSMAX) + ((2 * RGBMAX - cmax - cmin) // 2)) // (2 * RGBMAX - cmax - cmin)
    d = cmax - cmin
    Rd, Gd, Bd = ((((cmax - x) * (HLSMAX // 6)) + (d // 2)) // d for x in (R, G, B))
    if R == cmax:
        H = Bd - Gd
    elif G == cmax:
        H = (HLSMAX // 3) + Rd - Bd
    else:
        H = ((2 * HLSMAX) // 3) + Gd - Rd
    return H % HLSMAX if H < 0 or H > HLSMAX else H, L, S


def _hue(n1, n2, h):
    if h < 0:
        h += HLSMAX
    if h > HLSMAX:
        h -= HLSMAX
    if h < HLSMAX // 6:
        return n1 + (((n2 - n1) * h + (HLSMAX // 12)) // (HLSMAX // 6))
    if h < HLSMAX // 2:
        return n2
    if h < (HLSMAX * 2) // 3:
        return n1 + (((n2 - n1) * (((HLSMAX * 2) // 3) - h) + (HLSMAX // 12)) // (HLSMAX // 6))
    return n1


def _hls2rgb(H, L, S):
    if S == 0:
        v = (L * RGBMAX) // HLSMAX
        return v, v, v
    m2 = (L * (HLSMAX + S) + (HLSMAX // 2)) // HLSMAX if L <= HLSMAX // 2 else L + S - ((L * S) + (HLSMAX // 2)) // HLSMAX
    m1 = 2 * L - m2
    f = lambda h: min(255, max(0, (_hue(m1, m2, h) * RGBMAX + (HLSMAX // 2)) // HLSMAX))
    return f(H + HLSMAX // 3), f(H), f(H - HLSMAX // 3)


def ms_tint(rgb: str, tint: float) -> str:
    H, L, S = _rgb2hls(*(int(rgb[i:i + 2], 16) for i in (0, 2, 4)))
    l = L * (1 + tint) if tint < 0 else L * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint))
    return "%02X%02X%02X" % _hls2rgb(H, int(round(l)), S)


class Styles:
    def __init__(self, styles_xml: str, theme_xml: str):
        self.xml = styles_xml
        m = re.search(r'<fills count="\d+">(.*?)</fills>', styles_xml, re.S)
        self.fills = re.findall(r'<fill>.*?</fill>|<fill/>', m.group(1), re.S)
        m = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', styles_xml, re.S)
        self.xfs = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', m.group(1), re.S)
        idx = re.search(r'<indexedColors>(.*?)</indexedColors>', styles_xml, re.S)
        self.indexed = (re.findall(r'rgb="([0-9A-Fa-f]{8})"', idx.group(1)) if idx else None)
        self.theme = self._theme(theme_xml)
        self._fill_rgb = {}

    @staticmethod
    def _theme(theme_xml: str):
        m = re.search(r'<a:clrScheme[^>]*>(.*?)</a:clrScheme>', theme_xml, re.S)
        cols = {}
        for name, body in re.findall(r'<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>(.*?)</a:\1>', m.group(1), re.S):
            v = re.search(r'(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"', body)
            cols[name] = v.group(1).upper() if v else "000000"
        order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4",
                 "accent5", "accent6", "hlink", "folHlink"]
        return [cols.get(k, "000000") for k in order]

    def _color(self, tag: str | None):
        if not tag:
            return None
        a = dict(re.findall(r'(\w+)="([^"]*)"', tag))
        if "rgb" in a:
            rgb = a["rgb"][-6:].upper()
        elif "theme" in a:
            rgb = self.theme[int(a["theme"])]
        elif "indexed" in a:
            i = int(a["indexed"])
            pal = self.indexed or DEFAULT_INDEXED
            rgb = (pal[i][-6:] if i < len(pal) else "000000").upper()
        elif a.get("auto") == "1":
            rgb = "000000"
        else:
            return None
        tint = float(a.get("tint", 0) or 0)
        if tint:
            rgb = ms_tint(rgb, tint)
        return rgb

    def fill_id(self, s: int) -> int:
        return int(re.search(r'fillId="(\d+)"', self.xfs[s]).group(1))

    def interior_rgb(self, s: int | None) -> str:
        """Range.Interior.Color ekvivalents (RGB hex)."""
        if s is None:
            s = 0
        fid = self.fill_id(s)
        if fid in self._fill_rgb:
            return self._fill_rgb[fid]
        f = self.fills[fid]
        pt = re.search(r'patternType="(\w+)"', f)
        pt = pt.group(1) if pt else None
        fg = re.search(r'<fgColor[^>]*/>', f)
        bg = re.search(r'<bgColor[^>]*/>', f)
        if pt in (None, "none"):
            rgb = "FFFFFF"
        elif pt == "solid":
            rgb = self._color(fg.group(0) if fg else None) or "000000"
        else:
            rgb = self._color(bg.group(0) if bg else None) or "FFFFFF"
        self._fill_rgb[fid] = rgb
        return rgb

    # --- jauns/atrasts xf ar citu fillId (pārējais nemainās) ---
    @staticmethod
    def _xf_parts(xf: str):
        m = re.match(r'<xf ([^>]*?)\s*(/>|>(.*)</xf>)$', xf, re.S)
        return dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))), (m.group(3) or "")

    def xf_with_fill(self, s: int, fill_id: int) -> int:
        """Indekss xf, kas ir tieši tāds pats kā `s`, tikai ar citu fillId (atrod vai pievieno)."""
        attrs, kids = self._xf_parts(self.xfs[s])
        attrs["fillId"] = str(fill_id)
        if fill_id == 0:
            attrs.pop("applyFill", None)
        else:
            attrs["applyFill"] = "1"
        for i, x in enumerate(self.xfs):
            a, k = self._xf_parts(x)
            if a == attrs and k == kids:
                return i
        order = ["numFmtId", "fontId", "fillId", "borderId", "xfId", "quotePrefix", "pivotButton",
                 "applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection"]
        keys = [k for k in order if k in attrs] + [k for k in attrs if k not in order]
        head = "<xf " + " ".join(f'{k}="{attrs[k]}"' for k in keys)
        self.xfs.append(head + (f">{kids}</xf>" if kids else "/>"))
        return len(self.xfs) - 1

    def dump(self) -> str:
        body = "".join(self.xfs)
        return re.sub(r'<cellXfs count="\d+">.*?</cellXfs>',
                      lambda _: f'<cellXfs count="{len(self.xfs)}">{body}</cellXfs>', self.xml, count=1, flags=re.S)


# ----------------------------------------------------------------------------- shared strings
class SST:
    def __init__(self, xml: str):
        self.xml = xml
        self.items = re.findall(r'<si>(.*?)</si>', xml, re.S)
        self.text = [html.unescape("".join(re.findall(r'<t[^>]*>(.*?)</t>', it, re.S))) for it in self.items]
        self.added = []

    def index(self, s: str) -> int:
        for i, t in enumerate(self.text):
            if t == s and "<r>" not in self.items[i]:
                return i
        self.text.append(s)
        esc = html.escape(s, quote=False)
        self.items.append(f"<t>{esc}</t>")
        self.added.append(f"<si><t>{esc}</t></si>")
        return len(self.text) - 1

    def dump(self) -> str:
        if not self.added:
            return self.xml
        x = self.xml.replace("</sst>", "".join(self.added) + "</sst>")
        x = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{len(self.items)}"', x, count=1)
        m = re.search(r' count="(\d+)"', x)
        if m:
            x = x.replace(m.group(0), f' count="{int(m.group(1)) + len(self.added)}"', 1)
        return x


# ----------------------------------------------------------------------------- lapa
class Cell:
    __slots__ = ("r", "c", "attrs", "inner", "start", "end", "s", "t", "f", "fkind", "v", "dirty")

    def __init__(self, r, c, attrs, inner, start, end):
        self.r, self.c, self.attrs, self.inner, self.start, self.end = r, c, attrs, inner, start, end
        a = dict(re.findall(r'(\w+)="([^"]*)"', attrs))
        self.s = int(a["s"]) if "s" in a else None
        self.t = a.get("t")
        self.f = None
        self.fkind = None
        self.v = None
        self.dirty = False


class Sheet:
    def __init__(self, xml: str, sst: SST, styles: Styles):
        self.xml, self.sst, self.styles = xml, sst, styles
        self.cells: dict[tuple[int, int], Cell] = {}
        masters = {}
        children = []
        for m in CELL_RE.finditer(xml):
            col, row, attrs, inner = m.group(1), int(m.group(2)), m.group(3), m.group(4) or ""
            c = Cell(row, col2n(col), attrs, inner, m.start(), m.end())
            fm = re.search(r'<f([^>]*)>(.*?)</f>', inner, re.S) or re.search(r'<f([^>]*)/>', inner)
            if fm:
                fa = dict(re.findall(r'(\w+)="([^"]*)"', fm.group(1)))
                c.fkind = fa.get("t", "normal")
                text = html.unescape(fm.group(2)) if fm.lastindex >= 2 else ""
                if c.fkind == "shared":
                    if text:
                        masters[fa["si"]] = (row, c.c, text)
                        c.f = text
                    else:
                        children.append((c, fa["si"]))
                else:
                    c.f = text
            vm = re.search(r'<v>(.*?)</v>', inner, re.S)
            if c.t == "s" and vm:
                c.v = sst.text[int(vm.group(1))]
            elif c.t == "inlineStr":
                c.v = html.unescape("".join(re.findall(r'<t[^>]*>(.*?)</t>', inner, re.S)))
            elif c.t == "str":
                c.v = html.unescape(vm.group(1)) if vm else ""
            elif c.t == "e":
                c.v = XLError(vm.group(1) if vm else "#VALUE!")
            elif c.t == "b":
                c.v = bool(int(vm.group(1))) if vm else None
            elif vm:
                c.v = float(vm.group(1))
            self.cells[(row, c.c)] = c
        for c, si in children:
            mr, mc, text = masters[si]
            c.f = shift_formula(text, c.r - mr, c.c - mc)
        self.cached = {k: c.v for k, c in self.cells.items() if c.f is not None}
        self._memo = {}
        self.edits: dict[tuple[int, int], str] = {}  # (r,c) -> jauns <c> XML

    # --- vērtības ---
    def value(self, r, c):
        cell = self.cells.get((r, c))
        if cell is None:
            return None
        if cell.f is not None:
            key = (r, c)
            if key not in self._memo:
                self._memo[key] = None  # cikla aizsardzība
                try:
                    self._memo[key] = Evaluator(self, r, c).run(cell.f)
                except XLError as e:
                    self._memo[key] = e
            return self._memo[key]
        return cell.v

    def color(self, r, c):
        cell = self.cells.get((r, c))
        return self.styles.interior_rgb(cell.s if cell else None)

    def recalc(self):
        self._memo = {}
        return {k: self.value(*k) for k, c in self.cells.items() if c.f is not None}

    # --- rediģēšana (tikai konstantes) ---
    def set_cell(self, r, c, value, style: int):
        cell = self.cells.get((r, c))
        assert cell is not None, f"nav <c> elementa {n2col(c)}{r}"
        assert cell.f is None, f"{n2col(c)}{r} ir formula"
        ref = f"{n2col(c)}{r}"
        if value is None:
            x = f'<c r="{ref}" s="{style}"/>'
        elif isinstance(value, str):
            x = f'<c r="{ref}" s="{style}" t="s"><v>{self.sst.index(value)}</v></c>'
        else:
            x = f'<c r="{ref}" s="{style}"><v>{fmt_num(value)}</v></c>'
        cell.v, cell.s, cell.t = value, style, ("s" if isinstance(value, str) else None)
        self.edits[(r, c)] = x

    def dump(self, new_values: dict) -> tuple[str, list]:
        """Atgriež jauno sheet XML + sarakstu ar formulām, kuru vērtība mainījās."""
        changed = []
        repl = dict(self.edits)
        for k, nv in new_values.items():
            ov = self.cached.get(k)
            if same(ov, nv):
                continue
            cell = self.cells[k]
            changed.append((k, ov, nv))
            raw = self.xml[cell.start:cell.end]
            if isinstance(nv, XLError):
                code = str(nv)
                if code not in ("#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!", "#NAME?"):
                    raise RuntimeError(f"{n2col(k[1])}{k[0]}: formula dod nezināmu kļūdu {nv}")
                raw2 = re.sub(r'\s*t="(?:str|e|b)"', "", raw, count=1)
                raw2 = re.sub(r'^(<c [^>]*?)(\s*/?>)', r'\1 t="e"\2', raw2, count=1)
                newraw = re.sub(r'<v>.*?</v>', f"<v>{code}</v>", raw2, flags=re.S) if "<v>" in raw2 \
                    else raw2.replace("</c>", f"<v>{code}</v></c>")
                repl[k] = newraw
                continue
            if isinstance(nv, str):
                newraw = re.sub(r'<v>.*?</v>', f"<v>{html.escape(nv, quote=False)}</v>", raw, flags=re.S)
            else:
                if 't="str"' in raw or 't="e"' in raw:
                    raw = re.sub(r'\s*t="(?:str|e)"', "", raw, count=1)
                newraw = re.sub(r'<v>.*?</v>', f"<v>{fmt_num(nv if nv is not None else 0)}</v>", raw, flags=re.S) \
                    if "<v>" in raw else raw.replace("</c>", f"<v>{fmt_num(nv or 0)}</v></c>")
            repl[k] = newraw
        out, pos = [], 0
        for k in sorted(repl, key=lambda k: self.cells[k].start):
            cell = self.cells[k]
            out.append(self.xml[pos:cell.start])
            out.append(repl[k])
            pos = cell.end
        out.append(self.xml[pos:])
        return "".join(out), changed


def same(a, b):
    if isinstance(a, XLError) or isinstance(b, XLError):
        return str(a) == str(b)
    if a is None and b in (None, 0, 0.0, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
    return a == b


def shift_formula(text: str, dr: int, dc: int) -> str:
    parts = re.split(r'("(?:[^"]|"")*")', text)
    for i in range(0, len(parts), 2):
        def sh(m):
            ca, col, ra, row = m.groups()
            if m.start() > 0 and (parts[i][m.start() - 1].isalpha() or parts[i][m.start() - 1] == "_"):
                return m.group(0)
            c = col2n(col) + (0 if ca else dc)
            r = int(row) + (0 if ra else dr)
            return f"{ca}{n2col(c)}{ra}{r}"
        parts[i] = REF_RE.sub(sh, parts[i])
    return "".join(parts)


# ----------------------------------------------------------------------------- formulu novērtētājs
TOK = re.compile(r'\s*(?:(?P<err>#(?:REF!|N/A|NAME\?|VALUE!|DIV/0!|NUM!|NULL!))|(?P<str>"(?:[^"]|"")*")|(?P<range>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)'
                 r'|(?P<func>[A-Za-z_][A-Za-z0-9_.]*)\(|(?P<ref>\$?[A-Z]{1,3}\$?\d+)'
                 r'|(?P<num>\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|\.\d+)|(?P<op><>|<=|>=|[-+*/^&=<>(),]))')


class Rng:
    def __init__(self, r1, c1, r2, c2):
        self.r1, self.c1, self.r2, self.c2 = min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)

    def coords(self):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield r, c


def parse_ref(s):
    m = REF_RE.fullmatch(s)
    return int(m.group(4)), col2n(m.group(2))


class Evaluator:
    def __init__(self, sheet: Sheet, r, c):
        self.sh = sheet

    def run(self, f: str):
        f = f.lstrip("=")
        self.toks, pos = [], 0
        while pos < len(f):
            m = TOK.match(f, pos)
            if not m or m.end() == pos:
                if f[pos:].strip() == "":
                    break
                raise XLError(f"#PARSE({f[pos:pos + 10]})")
            kind = m.lastgroup
            self.toks.append((kind, m.group(kind)))
            pos = m.end()
        self.i = 0
        v = self.expr()
        if self.i != len(self.toks):
            raise XLError("#PARSE")
        return self.scalar(v)

    # --- palīgi ---
    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def scalar(self, v):
        if isinstance(v, Rng):
            if v.r1 == v.r2 and v.c1 == v.c2:
                return self.sh.value(v.r1, v.c1)
            raise XLError("#VALUE!")
        return v

    def num(self, v):
        v = self.scalar(v)
        if isinstance(v, XLError):
            raise v
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            raise XLError("#VALUE!")

    # --- gramatika ---
    def expr(self):
        a = self.concat()
        while self.peek() in (("op", "="), ("op", "<>"), ("op", "<"), ("op", ">"), ("op", "<="), ("op", ">=")):
            op = self.take()[1]
            b = self.concat()
            x, y = self.scalar(a), self.scalar(b)
            a = {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]
        return a

    def concat(self):
        a = self.add()
        while self.peek() == ("op", "&"):
            self.take()
            b = self.add()
            a = f"{self._s(a)}{self._s(b)}"
        return a

    def _s(self, v):
        v = self.scalar(v)
        return "" if v is None else (fmt_num(v) if isinstance(v, float) else str(v))

    def add(self):
        a = self.mul()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            b = self.mul()
            a = self.num(a) + self.num(b) if op == "+" else self.num(a) - self.num(b)
        return a

    def mul(self):
        a = self.pow()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            b = self.pow()
            if op == "*":
                a = self.num(a) * self.num(b)
            else:
                d = self.num(b)
                if d == 0:
                    raise XLError("#DIV/0!")
                a = self.num(a) / d
        return a

    def pow(self):
        a = self.unary()
        while self.peek() == ("op", "^"):
            self.take()
            a = self.num(a) ** self.num(self.unary())
        return a

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -self.num(self.unary())
        if self.peek() == ("op", "+"):
            self.take()
            return self.num(self.unary())
        return self.primary()

    def primary(self):
        kind, val = self.take()
        if kind == "err":
            raise XLError(val)
        if kind == "num":
            return float(val)
        if kind == "str":
            return val[1:-1].replace('""', '"')
        if kind == "ref":
            return Rng(*parse_ref(val), *parse_ref(val))
        if kind == "range":
            a, b = val.split(":")
            return Rng(*parse_ref(a), *parse_ref(b))
        if kind == "func":
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.expr())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.expr())
            assert self.take() == ("op", ")")
            return self.call(val.upper(), args)
        if (kind, val) == ("op", "("):
            v = self.expr()
            assert self.take() == ("op", ")")
            return v
        raise XLError(f"#PARSE({val})")

    # --- funkcijas ---
    def call(self, name, args):
        sh = self.sh
        if name == "SUM":
            t = 0.0
            for a in args:
                if isinstance(a, Rng):
                    for r, c in a.coords():
                        v = sh.value(r, c)
                        if isinstance(v, XLError):
                            raise v
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            t += v
                else:
                    t += self.num(a)
            return t
        if name == "SUMIF":
            rng, crit = args[0], self.scalar(args[1])
            srng = args[2] if len(args) > 2 else rng
            t = 0.0
            for r, c in rng.coords():
                v = sh.value(r, c)
                if isinstance(crit, str) and isinstance(v, str) and v.strip().lower() == crit.lower() \
                        or (not isinstance(crit, str) and v == crit):
                    sv = sh.value(srng.r1 + (r - rng.r1), srng.c1 + (c - rng.c1))
                    if isinstance(sv, (int, float)) and not isinstance(sv, bool):
                        t += sv
            return t
        if name == "SUMBYHEADERCOLOR":
            hdr, rng = args
            hc = sh.color(hdr.r1, hdr.c1)
            t = 0.0
            for r, c in rng.coords():
                v = sh.value(r, c)
                if v is None or isinstance(v, XLError):
                    continue
                if isinstance(v, bool):
                    nv = -1.0 if v else 0.0
                elif isinstance(v, (int, float)):
                    nv = float(v)
                else:
                    try:
                        nv = float(str(v).strip().replace(",", "."))
                    except ValueError:
                        continue
                if sh.color(r, c) == hc:
                    t += nv
            return t
        raise XLError(f"#NAME?({name})")


# ----------------------------------------------------------------------------- piezīmes (comments + VML)
class Notes:
    """Šūnu piezīmes: xl/comments1.xml + xl/drawings/vmlDrawing1.vml (Excel formātā)."""

    def __init__(self, comments_xml: str, vml: str):
        self.xml, self.vml = comments_xml, vml

    @staticmethod
    def _ref(r, c):
        return f"{n2col(c)}{r}"

    def refs(self):
        out = []
        for m in re.finditer(r'<comment ref="([A-Z]+)(\d+)"', self.xml):
            out.append((int(m.group(2)), col2n(m.group(1))))
        return out

    def drop(self, r, c) -> bool:
        self.xml, n = re.subn(r'<comment ref="%s" [^>]*>.*?</comment>' % self._ref(r, c), "", self.xml, flags=re.S)
        if n:
            pat = (r'<v:shape [^>]*>(?:(?!</v:shape>).)*?<x:Row>%d</x:Row>\s*<x:Column>%d</x:Column>.*?</v:shape>'
                   % (r - 1, c - 1))
            self.vml, m = re.subn(pat, "", self.vml, flags=re.S)
            assert m == 1, f"VML forma {self._ref(r, c)}: {m}"
        return bool(n)

    def drop_many(self, cells) -> int:
        cells = set(cells)
        keep, n = [], 0
        for m in re.finditer(r'<comment ref="([A-Z]+)(\d+)" [^>]*>.*?</comment>', self.xml, re.S):
            if (int(m.group(2)), col2n(m.group(1))) in cells:
                n += 1
            else:
                keep.append(m.group(0))
        self.xml = re.sub(r'<commentList>.*</commentList>', lambda _: "<commentList>" + "".join(keep) + "</commentList>",
                          self.xml, count=1, flags=re.S)
        rc = {(r - 1, c - 1) for r, c in cells}

        def shape(m):
            rr = re.search(r'<x:Row>(\d+)</x:Row>\s*<x:Column>(\d+)</x:Column>', m.group(0))
            return "" if rr and (int(rr.group(1)), int(rr.group(2))) in rc else m.group(0)
        self.vml = re.sub(r'<v:shape [^>]*>(?:(?!</v:shape>).)*</v:shape>', shape, self.vml, flags=re.S)
        return n

    def add(self, r, c, text):
        assert f'<comment ref="{self._ref(r, c)}"' not in self.xml
        uids = [int(u[-8:], 16) for u in re.findall(r'xr:uid="\{([0-9A-F-]+)\}"', self.xml)] or [0]
        uid = "{00000000-0006-0000-0000-%012X}" % (max(uids) + 1)
        x = (f'<comment ref="{self._ref(r, c)}" authorId="0" shapeId="0" xr:uid="{uid}"><text><r><rPr><sz val="11"/>'
             f'<color theme="1"/><rFont val="Calibri"/><scheme val="minor"/></rPr>'
             f'<t>{html.escape(text, quote=False)}</t></r></text></comment>')
        pos = next((m.start() for m in re.finditer(r'<comment ref="([A-Z]+)(\d+)"', self.xml)
                    if (int(m.group(2)), col2n(m.group(1))) > (r, c)), self.xml.index("</commentList>"))
        self.xml = self.xml[:pos] + x + self.xml[pos:]
        sid = max([int(i) for i in re.findall(r'_x0000_s(\d+)', self.vml)] or [1024]) + 1
        z = max([int(i) for i in re.findall(r'z-index:(\d+)', self.vml)] or [0]) + 1
        # lodziņa izmērs: ~60 zīmes rindā pie 360 pt (Calibri 11), garās rindas aplaužas
        r0, c0 = r - 1, c - 1
        lines = sum(max(1, math.ceil(len(t) / 60)) for t in text.split("\n"))
        self.vml = self.vml.replace("</xml>", (
            f'<v:shape id="_x0000_s{sid}" type="#_x0000_t202" style=\'position:absolute;\n'
            f'  margin-left:0pt;margin-top:0pt;width:360pt;height:{15 * lines + 2.25}pt;\n'
            f'  z-index:{z};visibility:hidden;mso-wrap-style:square\' fillcolor="infoBackground [80]"\n'
            f'  strokecolor="none [81]" o:insetmode="auto">\n  <v:fill color2="infoBackground [80]"/>\n'
            f'  <v:shadow color="none [81]" obscured="t"/>\n  <v:path o:connecttype="none"/>\n'
            f'  <v:textbox style=\'mso-direction-alt:auto;mso-fit-shape-to-text:t\'>\n'
            f'   <div style=\'text-align:left\'></div>\n  </v:textbox>\n'
            f'  <x:ClientData ObjectType="Note">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n'
            f'   <x:Anchor>\n    {c0 + 1}, 3, {max(0, r0 - 1)}, 20, {c0 + 13}, 10, {r0 + lines}, 14</x:Anchor>\n'
            f'   <x:AutoFill>False</x:AutoFill>\n   <x:Row>{r0}</x:Row>\n   <x:Column>{c0}</x:Column>\n'
            f'  </x:ClientData>\n </v:shape>') + "</xml>")


def day_groups(sh: "Sheet") -> dict:
    """Dienu grupas: sapludinātie 8. rindas apgabali ar vērtību 1..31 -> kolonnu saraksts."""
    groups = {}
    for a, b in re.findall(r'<mergeCell ref="([A-Z]+8):([A-Z]+8)"/>', sh.xml):
        c1, c2 = col2n(a[:-1]), col2n(b[:-1])
        v = sh.value(8, c1)
        if isinstance(v, float) and 1 <= v <= 31:
            groups[int(v)] = list(range(c1, c2 + 1))
    return groups


# ----------------------------------------------------------------------------- kolonnu ievietošana
def _remap_refs(text: str, fc) -> str:
    """Pārraksta šūnu atsauces formulā / diapazonā: kolonna c -> fc(c). Teksts pēdiņās netiek aiztikts."""
    parts = re.split(r'("(?:[^"]|"")*")', text)
    for i in range(0, len(parts), 2):
        seg = parts[i]

        def sh(m, seg=seg):
            if m.start() > 0 and (seg[m.start() - 1].isalnum() or seg[m.start() - 1] in "_."):
                return m.group(0)
            ca, col, ra, row = m.groups()
            return f"{ca}{n2col(fc(col2n(col)))}{ra}{row}"
        parts[i] = REF_RE.sub(sh, seg)
    return "".join(parts)


def insert_columns(parts: dict, inserts: list[tuple[int, int]], extend_merge_row: int = 8) -> dict:
    """Ievieto kolonnas lapā sheet1. inserts = [(pos, k)]: k jaunas kolonnas PIRMS kolonnas `pos`
    (1-bāzēta, vecā numerācija). Jaunās šūnas saņem kreisās kaimiņšūnas stilu; 8. rindas
    sapludinājums, kas beidzas tieši pirms `pos`, tiek pagarināts. Koplietotās formulas tiek
    pārvērstas parastās. calcChain tiek dzēsts (Excel to atjauno)."""
    inserts = sorted(inserts)

    def fc(c):
        return c + sum(k for p, k in inserts if c >= p)

    sheet = parts["xl/worksheets/sheet1.xml"]
    sst = SST(parts["xl/sharedStrings.xml"])
    styles = Styles(parts["xl/styles.xml"], parts["xl/theme/theme1.xml"])
    sh = Sheet(sheet, sst, styles)

    # --- sheetData ---
    def cell_xml(c: Cell, newref: str) -> str:
        attrs = re.sub(r'\s*r="[^"]*"', "", c.attrs)
        inner = c.inner or ""
        if c.f is not None:
            fm = re.search(r'<f([^>]*)>(.*?)</f>', inner, re.S) or re.search(r'<f([^>]*)/>', inner)
            fa = dict(re.findall(r'(\w+)="([^"]*)"', fm.group(1)))
            keep = []
            if fa.get("t") == "array":
                a, b = (fa["ref"].split(":") + [fa["ref"]])[:2]
                keep.append(f't="array" ref="{_remap_refs(fa["ref"], fc)}"')
            if fa.get("ca") == "1":
                keep.append('ca="1"')
            ftxt = html.escape(_remap_refs(c.f, fc), quote=False)
            newf = f'<f{(" " + " ".join(keep)) if keep else ""}>{ftxt}</f>'
            inner = inner[:fm.start()] + newf + inner[fm.end():]
        return f'<c r="{newref}"{attrs}>{inner}</c>' if inner else f'<c r="{newref}"{attrs}/>'

    by_row: dict[int, list[Cell]] = {}
    for (r, c), cell in sh.cells.items():
        by_row.setdefault(r, []).append(cell)

    def row_repl(m):
        rattrs, body = m.group(1), m.group(2)
        r = int(re.search(r'\br="(\d+)"', rattrs).group(1))
        rattrs = re.sub(r'\s*spans="[^"]*"', "", rattrs)
        cells = {cell.c: cell for cell in by_row.get(r, [])}
        out = []
        for c in sorted(cells):
            out.append((fc(c), cell_xml(cells[c], f"{n2col(fc(c))}{r}")))
        for p, k in inserts:
            left = cells.get(p - 1)
            if left is not None:
                for j in range(k):
                    nc = fc(p - 1) + 1 + j
                    out.append((nc, f'<c r="{n2col(nc)}{r}"{(" s=" + chr(34) + str(left.s) + chr(34)) if left.s is not None else ""}/>'))
        out.sort()
        return f"<row{rattrs}>" + "".join(x for _, x in out) + "</row>"

    head, rest = sheet.split("<sheetData>", 1)
    data, tail = rest.split("</sheetData>", 1)
    data = re.sub(r'<row((?:[^>/]|/(?!>))*)>(.*?)</row>', row_repl, data, flags=re.S)
    data = re.sub(r'<row([^>]*?)/>', lambda m: f"<row{re.sub(r' spans=.[^ ]*.', '', m.group(1))}/>", data)

    # --- head: dimension, selection, pane, cols ---
    head = re.sub(r'(<dimension ref=")([^"]+)"', lambda m: m.group(1) + _remap_refs(m.group(2), fc) + '"', head)
    head = re.sub(r'((?:activeCell|topLeftCell|sqref)=")([^"]+)"', lambda m: m.group(1) + _remap_refs(m.group(2), fc) + '"', head)
    cols = re.findall(r'<col [^>]*/>', head)
    newcols = []
    for x in cols:
        a = dict(re.findall(r'(\w+)="([^"]*)"', x))
        mn, mx = int(a["min"]), int(a["max"])
        nmn = fc(mn)
        nmx = fc(mx)  # ja ievietošana ir starp min un max, diapazons izstiepjas un pārklāj jaunās kolonnas
        x2 = re.sub(r'min="\d+"', f'min="{nmn}"', re.sub(r'max="\d+"', f'max="{nmx}"', x))
        newcols.append((nmn, nmx, x2))
    covered = set()
    for mn, mx, _ in newcols:
        covered.update(range(mn, mx + 1))
    for p, k in inserts:
        for j in range(k):
            nc = fc(p - 1) + 1 + j
            if nc not in covered:
                src = next((x for mn, mx, x in newcols if mn <= fc(p - 1) <= mx), None)
                if src:
                    newcols.append((nc, nc, re.sub(r'min="\d+"', f'min="{nc}"', re.sub(r'max="\d+"', f'max="{nc}"', src))))
    newcols.sort()
    head = re.sub(r'<cols>.*</cols>', lambda _: "<cols>" + "".join(x for _, _, x in newcols) + "</cols>", head, flags=re.S)

    # --- tail: merges, CF, colBreaks ---
    def merge_repl(m):
        ref = _remap_refs(m.group(1), fc)
        a, b = ref.split(":")
        (ac, ar), (bc, br) = re.match(r'([A-Z]+)(\d+)', a).groups(), re.match(r'([A-Z]+)(\d+)', b).groups()
        if int(ar) == int(br) == extend_merge_row:
            old_end = None
            for p, k in inserts:
                if col2n(bc) == fc(p - 1):
                    bc = n2col(col2n(bc) + k)
        return f'<mergeCell ref="{ac}{ar}:{bc}{br}"/>'
    tail = re.sub(r'<mergeCell ref="([^"]+)"/>', merge_repl, tail)
    tail = re.sub(r'(<conditionalFormatting sqref=")([^"]+)"', lambda m: m.group(1) + _remap_refs(m.group(2), fc) + '"', tail)
    tail = re.sub(r'(<formula>)(.*?)(</formula>)', lambda m: m.group(1) + html.escape(_remap_refs(html.unescape(m.group(2)), fc), quote=False) + m.group(3), tail)
    tail = re.sub(r'(<brk id=")(\d+)"', lambda m: m.group(1) + str(fc(int(m.group(2)))) + '"', tail)
    out = dict(parts)
    out["xl/worksheets/sheet1.xml"] = head + "<sheetData>" + data + "</sheetData>" + tail

    # --- piezīmes ---
    out["xl/comments1.xml"] = re.sub(r'(<comment ref=")([A-Z]+\d+)"', lambda m: m.group(1) + _remap_refs(m.group(2), fc) + '"', parts["xl/comments1.xml"])

    def vml_shape(m):
        s = m.group(0)
        s = re.sub(r'<x:Column>(\d+)</x:Column>', lambda q: f"<x:Column>{fc(int(q.group(1)) + 1) - 1}</x:Column>", s)

        def anc(q):
            v = [int(t) for t in q.group(1).split(",")]
            v[0] = fc(v[0] + 1) - 1
            v[4] = fc(v[4] + 1) - 1
            return "<x:Anchor>\n    " + ", ".join(str(t) for t in v) + "</x:Anchor>"
        return re.sub(r'<x:Anchor>\s*([\d,\s]+?)</x:Anchor>', anc, s)
    out["xl/drawings/vmlDrawing1.vml"] = re.sub(r'<v:shape [^>]*>(?:(?!</v:shape>).)*</v:shape>', vml_shape,
                                                parts["xl/drawings/vmlDrawing1.vml"], flags=re.S)

    # --- calcChain dzēšana ---
    out.pop("xl/calcChain.xml", None)
    out["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", parts["[Content_Types].xml"])
    out["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", parts["xl/_rels/workbook.xml.rels"])
    return out
