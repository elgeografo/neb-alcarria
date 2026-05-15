#!/usr/bin/env python3
"""
Pre-compute and write cached values for ALL formula cells in the I.B. budget
sheet, so the user sees correct totals immediately without depending on
Excel's recalculation (which is blocked by sheet protection + stale calcChain).

Formulas are preserved. Inputs and cached values are written in one pass.
"""

import re
import shutil
import zipfile
from pathlib import Path

V03 = Path(__file__).resolve().parent.parent / \
    'docu a presentar/03_18_mayo/V03/Anexo_I_SOLICITUD_PRESUP_Form.xlsm'
V04 = Path(__file__).resolve().parent.parent / \
    'docu a presentar/03_18_mayo/V04/Anexo_I_SOLICITUD_PRESUP_Form.xlsm'

TARGETS = [V04]  # only V04 is the active version


# (row, M1=COPRODELI P1, M1 P2, M2=UPM P1, M2 P2)
INPUTS: dict[int, tuple[float, float, float, float]] = {
    15: (0.00,         0.00,         0.00,         0.00),     # A.1
    16: (0.00,      7650.00,         0.00,      3690.00),     # A.2
    17: (2238.40,   4197.00,      2798.00,      6211.56),     # A.3
    18: (3059.56,   6716.87,      7505.96,     10764.85),     # A.4
    19: (6267.52,   6995.00,      6267.52,      6995.00),     # A.5
    21: (5427.78,  25042.27,      8895.42,     24353.93),     # B.1
    22: (500.00,       0.00,       500.00,         0.00),     # B.2
    23: (630.00,    1050.00,       420.00,       700.00),     # B.3
    24: (200.00,     300.00,       200.00,       300.00),     # B.4
    25: (0.00,      1400.00,         0.00,      1400.00),     # B.5
    26: (750.00,    1750.00,       750.00,      1750.00),     # B.6
    27: (250.00,     499.99,       250.00,       499.99),     # B.7
    29: (0.00,         0.00,      2833.63,      4048.04),     # C.1
    30: (0.00,     21700.00,         0.00,     14300.00),     # C.2
}

A_ROWS = [15, 16, 17, 18, 19]
B_ROWS = [21, 22, 23, 24, 25, 26, 27]
C_ROWS = [29, 30]


def round2(x: float) -> float:
    return round(x, 2)


def compute_values() -> dict[str, float]:
    """Compute the cached value for every formula cell in rows 15..39."""
    v: dict[str, float] = {}

    # Inputs (G, H, J, K)
    for r, (g, h, j, k) in INPUTS.items():
        v[f'G{r}'], v[f'H{r}'], v[f'J{r}'], v[f'K{r}'] = g, h, j, k
        # Per-partida formulas
        v[f'I{r}'] = round2(g + h)          # M1 TOTAL = G + H
        v[f'L{r}'] = round2(j + k)          # M2 TOTAL = J + K
        v[f'E{r}'] = round2(g + j)          # P1 = sum across members (only 2)
        v[f'F{r}'] = round2(h + k)          # P2 = sum across members
        v[f'D{r}'] = round2(v[f'E{r}'] + v[f'F{r}'])  # TOTAL row

    # Row 20: A. PERSONAL = sum rows 15..19 (cell-by-cell)
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}20'] = round2(sum(v.get(f'{col}{r}', 0) for r in A_ROWS))
    # Row 28: B. SERVICIOS = sum rows 21..27
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}28'] = round2(sum(v.get(f'{col}{r}', 0) for r in B_ROWS))
    # Row 31: C. INVERSIONES = sum rows 29..30
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}31'] = round2(sum(v.get(f'{col}{r}', 0) for r in C_ROWS))

    # Row 32: TOTAL COSTES DIRECTOS (A+B+C) = row20 + row28 + row31
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}32'] = round2(v[f'{col}20'] + v[f'{col}28'] + v[f'{col}31'])

    # Row 33: TOTAL COSTES INDIRECTOS = 15% × A. PERSONAL (row 20)
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}33'] = round2(0.15 * v[f'{col}20'])

    # Row 34: SUBTOTAL SUBV. BIENES INVENTARIABLES = 100% × C.1 (row 29)
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}34'] = round2(v[f'{col}29'])

    # Row 35: TOTAL SUBVENCIÓN = A + B + C.2 + INDIRECTOS + SUBV.C.1
    #         = row20 + row28 + row30 + row33 + row34
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}35'] = round2(
            v[f'{col}20'] + v[f'{col}28'] + v[f'{col}30']
            + v[f'{col}33'] + v[f'{col}34']
        )

    # Row 36: SUBTOTAL FONDOS PROPIOS INVERSIONES = C.1 − SUBV.C.1 = 0
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}36'] = round2(v[f'{col}29'] - v[f'{col}34'])

    # Row 37: SUBTOTAL FONDOS PROPIOS = manual input, leave at 0
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}37'] = 0.0

    # Row 38: TOTAL FONDOS PROPIOS = row36 + row37
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}38'] = round2(v[f'{col}36'] + v[f'{col}37'])

    # Row 39: TOTAL PRESUPUESTO = TOTAL SUBV + TOTAL FONDOS PROPIOS
    for col in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'):
        v[f'{col}39'] = round2(v[f'{col}35'] + v[f'{col}38'])

    return v


def get_style(cell_xml: str) -> str | None:
    m = re.search(r's="([^"]+)"', cell_xml)
    return m.group(1) if m else None


def patch_sheet(xml: str, values: dict[str, float]) -> str:
    """Write cached <v> for every cell in `values`, preserving the formula
    (if any) and the style."""
    for coord, val in values.items():
        pat = re.compile(
            rf'<c r="{re.escape(coord)}"(?P<attrs>[^/>]*)(?:/>|>(?P<body>.*?)</c>)',
            re.DOTALL,
        )
        m = pat.search(xml)
        if not m:
            raise RuntimeError(f'cell {coord} not found')

        attrs = m.group('attrs') or ''
        body = m.group('body') or ''

        # Preserve any <f>...</f> (formula or shared-formula ref) in body
        fm = re.search(r'<f[^/]*(?:/>|>.*?</f>)', body, re.DOTALL)
        formula = fm.group(0) if fm else ''

        # Strip existing t="..." (we want numeric default)
        attrs = re.sub(r'\s+t="[^"]+"', '', attrs)

        new_body = f'{formula}<v>{val:.2f}</v>'
        new_cell = f'<c r="{coord}"{attrs}>{new_body}</c>'

        xml = xml[: m.start()] + new_cell + xml[m.end():]
    return xml


def fix_title_formula(xml: str) -> str:
    """D8 title formula references F417 (I.C.1) — change to F368 (I.B.)."""
    return xml.replace(
        "<f>'DATOS GENERALES'!F417</f>",
        "<f>'DATOS GENERALES'!F368</f>",
        1,
    )


def unhide_row_622(sheet2_xml: str) -> str:
    """Unhide línea estratégica I.B. row in DATOS GENERALES."""
    return re.sub(
        r'(<row r="622"[^>]*) hidden="1"',
        r'\1',
        sheet2_xml,
        count=1,
    )


def force_recalc_workbook(wb_xml: str) -> str:
    if 'fullCalcOnLoad' not in wb_xml:
        wb_xml = re.sub(
            r'<calcPr([^/]*)/>',
            r'<calcPr\1 fullCalcOnLoad="1"/>',
            wb_xml,
            count=1,
        )
    return wb_xml


def process(path: Path) -> None:
    with zipfile.ZipFile(path, 'r') as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    values = compute_values()

    sheet8 = contents['xl/worksheets/sheet8.xml'].decode('utf-8')
    sheet8 = patch_sheet(sheet8, values)
    sheet8 = fix_title_formula(sheet8)
    contents['xl/worksheets/sheet8.xml'] = sheet8.encode('utf-8')

    sheet2 = contents['xl/worksheets/sheet2.xml'].decode('utf-8')
    sheet2 = unhide_row_622(sheet2)
    contents['xl/worksheets/sheet2.xml'] = sheet2.encode('utf-8')

    contents['xl/workbook.xml'] = force_recalc_workbook(
        contents['xl/workbook.xml'].decode('utf-8')
    ).encode('utf-8')

    # Drop calcChain.xml if present (force rebuild)
    if 'xl/calcChain.xml' in names:
        names = [n for n in names if n != 'xl/calcChain.xml']
        del contents['xl/calcChain.xml']
        ct = contents['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override[^/]*PartName="/xl/calcChain.xml"[^/]*/>', '', ct)
        contents['[Content_Types].xml'] = ct.encode('utf-8')
        rels = contents['xl/_rels/workbook.xml.rels'].decode('utf-8')
        rels = re.sub(r'<Relationship[^/]*Target="calcChain.xml"[^/]*/>', '', rels)
        contents['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

    tmp = path.with_suffix(path.suffix + '.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    shutil.move(tmp, path)
    print(f'OK · {path}')

    # Summary
    print(f'\n  D20  A. PERSONAL         {values["D20"]:>13,.2f} €')
    print(f'  D28  B. SERVICIOS        {values["D28"]:>13,.2f} €')
    print(f'  D31  C. INVERSIONES      {values["D31"]:>13,.2f} €')
    print(f'  D32  COSTES DIRECTOS     {values["D32"]:>13,.2f} €')
    print(f'  D33  INDIRECTOS (15%)    {values["D33"]:>13,.2f} €')
    print(f'  D34  SUBV. C.1           {values["D34"]:>13,.2f} €')
    print(f'  D35  TOTAL SUBVENCIÓN    {values["D35"]:>13,.2f} €')
    print(f'  D39  TOTAL PRESUPUESTO   {values["D39"]:>13,.2f} €')


def main() -> None:
    for p in TARGETS:
        process(p)


if __name__ == '__main__':
    main()
