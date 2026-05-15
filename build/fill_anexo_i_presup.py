#!/usr/bin/env python3
"""
Fill the "PRESUPUESTO P.I.B-Agrupación" sheet of the Anexo I FEGA form,
sourcing data from the master file `NEB-ALCARRIA-PRESUPUESTO.xlsx`
(desglosado a partir de Anexo XXII Memo Presupuesto).

Only the per-partida input cells (G, H, J, K per row) are written; all
totals (rows 20, 28, 31, 32, 33, 34, 35, 36, 38, 39, 41) auto-compute via
the form's formulas.

Also fixes a template bug: D8 ("Título de la Propuesta") references
DATOS GENERALES!F417 (I.C.1 title) when it should reference F368 (I.B.).

Member mapping in the Anexo I form (different from Anexo XXII):
  MIEMBRO 1 = Fundación COPRODELI  (representante de la agrupación)
  MIEMBRO 2 = Universidad Politécnica de Madrid (UPM)
"""

import re
import shutil
import zipfile
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / \
    'docu a presentar/03_18_mayo/V04/Anexo_I_SOLICITUD_PRESUP_Form.xlsm'
DST = SRC


# (row, partida, M1=COPRODELI P1, M1 P2, M2=UPM P1, M2 P2)
BUDGET: list[tuple[int, str, float, float, float, float]] = [
    # PERSONAL
    (15, 'A.1 Personal docente externo',         0.00,         0.00,         0.00,         0.00),
    (16, 'A.2 Personal de apoyo externo',        0.00,      7650.00,         0.00,      3690.00),
    (17, 'A.3 Personal docente propio',       2238.40,      4197.00,      2798.00,      6211.56),
    (18, 'A.4 Personal de apoyo propio',      3059.56,      6716.87,      7505.96,     10764.85),
    (19, 'A.5 Personal de coordinación',      6267.52,      6995.00,      6267.52,      6995.00),
    # SERVICIOS Y FUNGIBLES
    (21, 'B.1 Transporte, alojamiento y manutención',  5427.78,     25042.27,      8895.42,     24353.93),
    (22, 'B.2 Entornos virtuales',             500.00,         0.00,       500.00,         0.00),
    (23, 'B.3 Material didáctico y fungibles', 630.00,      1050.00,       420.00,       700.00),
    (24, 'B.4 Seguros',                        200.00,       300.00,       200.00,       300.00),
    (25, 'B.5 Auditoría',                        0.00,      1400.00,         0.00,      1400.00),
    (26, 'B.6 Alquileres',                     750.00,      1750.00,       750.00,      1750.00),
    (27, 'B.7 Costes de comunicación',         250.00,       499.99,       250.00,       499.99),
    # INVERSIONES
    (29, 'C.1 Bienes inventariables',            0.00,         0.00,      2833.63,      4048.04),
    (30, 'C.2 Otras inversiones',                0.00,     21700.00,         0.00,     14300.00),
]


def patch_cell(row_xml: str, coord: str, new_cell: str) -> str:
    """Replace cell `coord` in row_xml with `new_cell`."""
    pat = re.compile(
        rf'<c r="{re.escape(coord)}"(?P<a>[^/>]*)(?:/>|>(?P<b>.*?)</c>)',
        re.DOTALL,
    )
    m = pat.search(row_xml)
    if not m:
        raise RuntimeError(f'cell {coord} not found')
    return row_xml[: m.start()] + new_cell + row_xml[m.end():]


def get_style(row_xml: str, coord: str) -> str | None:
    m = re.search(rf'<c r="{re.escape(coord)}"([^/>]*)', row_xml)
    if not m:
        return None
    sm = re.search(r's="([^"]+)"', m.group(1))
    return sm.group(1) if sm else None


def num_cell(coord: str, value: float, style: str | None) -> str:
    s = f' s="{style}"' if style else ''
    return f'<c r="{coord}"{s}><v>{value:.2f}</v></c>'


def update_sheet9(xml: str) -> str:
    # 1) For each partida row, set G, H, J, K with the four numeric values.
    for row, _label, c_p1, c_p2, u_p1, u_p2 in BUDGET:
        pat = re.compile(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', re.DOTALL)
        m = pat.search(xml)
        if not m:
            raise RuntimeError(f'row {row} not found')
        body = m.group(2)
        for coord, val in (
            (f'G{row}', c_p1),
            (f'H{row}', c_p2),
            (f'J{row}', u_p1),
            (f'K{row}', u_p2),
        ):
            style = get_style(body, coord)
            body = patch_cell(body, coord, num_cell(coord, val, style))
        xml = xml[: m.start()] + m.group(1) + body + m.group(3) + xml[m.end():]

    # 2) Fix template bug: D8 title formula references F417 (I.C.1) but
    #    this is the I.B. sheet, so it should reference F368.
    xml = xml.replace(
        "<f>'DATOS GENERALES'!F417</f>",
        "<f>'DATOS GENERALES'!F368</f>",
        1,
    )
    return xml


def main() -> None:
    with zipfile.ZipFile(SRC, 'r') as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    # PRESUPUESTO P.I.B-Agrupación = sheet8.xml (per workbook.xml rId8)
    sheet8 = contents['xl/worksheets/sheet8.xml'].decode('utf-8')
    contents['xl/worksheets/sheet8.xml'] = update_sheet9(sheet8).encode('utf-8')

    # Also unhide row 622 (Línea estratégica I.B.) in DATOS GENERALES if Excel
    # has re-hidden it via VBA — the user can't see the input otherwise.
    sheet2 = contents['xl/worksheets/sheet2.xml'].decode('utf-8')
    sheet2 = re.sub(
        r'(<row r="622"[^>]*) hidden="1"',
        r'\1',
        sheet2,
        count=1,
    )
    contents['xl/worksheets/sheet2.xml'] = sheet2.encode('utf-8')

    tmp = DST.with_suffix(DST.suffix + '.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    shutil.move(tmp, DST)
    print(f'OK · {DST}')


if __name__ == '__main__':
    main()
