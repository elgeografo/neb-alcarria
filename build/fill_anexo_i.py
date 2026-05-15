#!/usr/bin/env python3
"""
Fill the Anexo I FEGA form (xlsm) with ALMUNIA 4.0 / NEB-ALCARRIA data.

Edits xl/worksheets/sheet2.xml (DATOS GENERALES) directly, preserving:
  - all VBA macros, defined names, data validations, styles, merges
  - the 3 inputs the user already entered (Tipo solicitante, Tipo solicitud,
    casilla I.B.)

Inputs are written as inline strings (t="inlineStr") or numbers, on the
correct merged cell anchors. Rows that the form's VBA had hidden (because
the user had not yet cascaded the "Agrupación" → "2 miembros" → "I.B." path)
are un-hidden so the user can see and edit the values in Excel.
"""

import re
import shutil
import zipfile
from html import escape
from pathlib import Path

SRC = Path('docu a presentar/03_18_mayo/V03/Anexo_I_SOLICITUD_PRESUP_Form.xlsm')
DST = Path('docu a presentar/03_18_mayo/V03/Anexo_I_SOLICITUD_PRESUP_Form.xlsm')

# (cell, value, kind)  — kind is 'str' (inline string) or 'num' (number)
UPDATES: list[tuple[str, str, str]] = [
    # ── 1.2 Datos de la agrupación ────────────────────────────────────────
    ('E69',  'ALMUNIA 4.0',                              'str'),  # Nombre agrupación
    ('I73',  '2',                                         'num'),  # Nº miembros

    # ── MIEMBRO 1 — Fundación COPRODELI (Representante) ─────────────────
    ('E86',  'Juridica',                                  'str'),  # Personalidad
    ('E88',  'Fundación COPRODELI',                       'str'),  # Razón social
    ('E90',  'G81984270',                                 'str'),  # NIF entidad
    # Representante legal de COPRODELI = Manuel Alegre Carvajal (apoderado único agrupación)
    ('E94',  '03108568A',                                 'str'),  # NIF rep. legal
    ('K94',  'Manuel',                                    'str'),  # Nombre
    ('E96',  'Alegre',                                    'str'),  # Primer apellido
    ('K96',  'Carvajal',                                  'str'),  # Segundo apellido
    # Domicilio COPRODELI
    ('E100', 'Calle Guadiana, 54',                        'str'),  # Calle/Plaza
    ('E102', 'Camarma de Esteruelas',                     'str'),  # Ciudad
    ('J102', '28816',                                     'str'),  # CP
    ('M102', 'España',                                    'str'),  # País
    ('E104', 'Comunidad de Madrid',                       'str'),  # CCAA
    ('I104', 'Madrid',                                    'str'),  # Provincia
    ('M104', '648284827',                                 'str'),  # Teléfono
    ('E106', 'gges@coprodeli.edu.pe',                     'str'),  # Email
    ('K106', 'www.alcarriaencantada.com',                 'str'),  # Web
    # Persona de contacto = misma que representante
    ('E110', 'Manuel',                                    'str'),  # Nombre contacto
    ('K110', 'Alegre',                                    'str'),  # Primer apellido
    ('E112', 'gges@coprodeli.edu.pe',                     'str'),  # Email contacto
    ('K112', '648284827',                                 'str'),  # Teléfono contacto
    # Tipo de entidad
    ('E116', 'Privada',                                   'str'),  # Pública/Privada
    ('E118', 'No',                                        'str'),  # ¿Exención IVA?
    ('E120', '14. Fundación o Asociación sin ánimo de lucro de ámbito nacional', 'str'),
    # Alcance territorial
    ('E126', 'Sí',                                        'str'),  # ¿Ámbito nacional?
    # Ámbito sectorial
    ('E133', 'Sector agroalimentario',                    'str'),

    # ── MIEMBRO 2 — Universidad Politécnica de Madrid (UPM) ─────────────
    ('E137', 'Juridica',                                  'str'),
    ('E139', 'Universidad Politécnica de Madrid',         'str'),
    ('E141', 'Q2818015F',                                 'str'),  # CIF UPM
    # Representante legal UPM (Rector u órgano con firma delegada) — pendiente
    # de confirmación nominal por el equipo UPM; se deja en blanco.
    # Domicilio Rectorado UPM
    ('E151', 'Calle Ramiro de Maeztu, 7',                 'str'),
    ('E153', 'Madrid',                                    'str'),
    ('J153', '28040',                                     'str'),
    ('M153', 'España',                                    'str'),
    ('E155', 'Comunidad de Madrid',                       'str'),
    ('I155', 'Madrid',                                    'str'),
    ('M155', '910670000',                                 'str'),
    ('K157', 'www.upm.es',                                'str'),
    # Tipo de entidad
    ('E161', 'Publica',                                   'str'),
    ('E163', 'Si',                                        'str'),  # Sí exención IVA
    ('E165', '11. Universidad o Instituto de Investigación Público', 'str'),
    ('E171', 'Sí',                                        'str'),  # Ámbito nacional
    ('E178', 'Sector agroalimentario',                    'str'),

    # ── 2.2 Programa I.B. (NEB-ALCARRIA) ─────────────────────────────────
    ('F368',
     'NEB-ALCARRIA — Living Lab digital agroforestal en la Alcarria '
     '(Pastrana, Guadalajara): cuaderno COPRODELI-IA + integración '
     'CUE/SIEX/SIGPAC, sensórica IoT, dronística RGB/multiespectral y '
     'co-construcción de domo geodésico NEB.',
     'str'),
    ('F370',
     'Demostración en condiciones reales sobre 3.000 m² cedidos por '
     'Pastrana de que la horticultura digital de microexplotación es '
     'viable, rentable, sostenible y bella (filosofía NEB). Validación '
     'TRL 7→8 del cuaderno COPRODELI-IA con apoyo decisional por IA, '
     'integrado con CUE/SIGPAC, stack open-source (QGIS/QField, WebODM, '
     'LoRa, Sentinel-2) y co-construcción de un domo geodésico, en 7 '
     'Living Labs (10 ediciones) sobre 2 ciclos productivos.',
     'str'),
    ('F372',
     'Frenar la despoblación y fomentar el asentamiento de jóvenes '
     'profesionales cualificados en la Alcarria, demostrando '
     'empíricamente que la microexplotación hortícola digitalizada con '
     'IA, IoT y dronística es una vía profesional viable, rentable y '
     'sostenible alineada con la Nueva Bauhaus Europea.',
     'str'),
    ('F374',
     'Hortalizas y frutos de huerta morisca alcarreña (Anexo I TFUE); '
     'productos forestales asociados (corcho aplicado al domo NEB).',
     'str'),
    ('F376',
     'Productos incluidos en el Anexo I del TFUE y productos forestales fuera del ANEXO I TFUE',
     'str'),

    # Objetivos específicos (casillas booleanas a–i)
    # Marcamos: b) competitividad+digitalización · e) gestión sostenible
    # de recursos · g) jóvenes agricultores · h) empleo/crecimiento rural
    ('AH381', '1', 'bool'),   # b) competitividad / digitalización
    ('AH384', '1', 'bool'),   # e) recursos naturales (agua, suelo)
    ('AH386', '1', 'bool'),   # g) jóvenes agricultores + desarrollo rural
    ('AH387', '1', 'bool'),   # h) empleo, igualdad de género, biodemografía rural

    # Áreas temáticas
    ('D392', '6. Otros / Desarrollo rural',               'str'),
    ('K392', '6.1. Digitalización de parcelas',           'str'),
    ('D396', '8. Agricultura',                            'str'),  # secundaria (BP3)
    ('K396', '8.4. Gestión de explotaciones',             'str'),

    # CCAA + provincias donde se ejecuta
    ('E400', 'Castilla La Mancha',                        'str'),
    ('K400', 'Guadalajara',                               'str'),
    ('E401', 'Comunidad de Madrid',                       'str'),
    ('K401', 'Madrid',                                    'str'),

    # Duración estimada de la propuesta (meses): nov 2026 → mar 2028 = 17 m
    ('H405', '17',                                        'num'),

    # ── 4. Líneas estratégicas ──────────────────────────────────────────
    # Programa temático IB — únicas seleccionadas; se prioriza la línea A
    # (cuaderno digital de explotación) por ser el núcleo de la propuesta.
    ('E622',
     'A). Propuestas que aborden la adaptación al uso del cuaderno digital de explotación',
     'str'),
]

# Rows that need to be made visible (the form's VBA had them hidden because
# the user had not yet cascaded all selections).
UNHIDE_ROWS = (
    # Datos de la agrupación
    list(range(67, 83))
    # Miembro 1 (COPRODELI)
    + list(range(84, 134))
    # Miembro 2 (UPM)
    + list(range(135, 179))
    # Sección 4 — encabezados de líneas estratégicas + I.B. seleccionada
    + [615, 616, 617, 618, 619, 620, 621, 622, 623]
    # CCAA secundaria (segundo registro)
    + [401]
)


def make_inline_cell(coord: str, value: str, style: str | None) -> str:
    safe = escape(value, quote=True)
    style_attr = f' s="{style}"' if style else ''
    return f'<c r="{coord}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{safe}</t></is></c>'


def make_num_cell(coord: str, value: str, style: str | None) -> str:
    style_attr = f' s="{style}"' if style else ''
    return f'<c r="{coord}"{style_attr}><v>{value}</v></c>'


def make_bool_cell(coord: str, value: str, style: str | None) -> str:
    style_attr = f' s="{style}"' if style else ''
    return f'<c r="{coord}"{style_attr} t="b"><v>{value}</v></c>'


def update_sheet(xml: str) -> str:
    # Group updates by row
    by_row: dict[int, list[tuple[str, str, str]]] = {}
    for coord, value, kind in UPDATES:
        m = re.match(r'^([A-Z]+)(\d+)$', coord)
        if not m:
            raise ValueError(f'bad coord {coord}')
        row = int(m.group(2))
        by_row.setdefault(row, []).append((coord, value, kind))

    rows_touched = set()

    def patch_row(match: re.Match) -> str:
        row_open = match.group(1)
        row_attrs = match.group(2)
        row_body = match.group(3)
        row_close = match.group(4)
        row_num = int(re.search(r'r="(\d+)"', row_attrs).group(1))

        # Unhide if requested
        new_attrs = row_attrs
        if row_num in UNHIDE_ROWS and 'hidden="1"' in new_attrs:
            new_attrs = new_attrs.replace(' hidden="1"', '')

        # Apply updates for this row, if any
        if row_num in by_row:
            for coord, value, kind in by_row[row_num]:
                # Find existing <c r="COORD" .../> (self-closing) or with content
                cell_pat = re.compile(
                    rf'<c r="{re.escape(coord)}"(?P<a>[^/>]*)(?:/>|>(?P<b>.*?)</c>)',
                    re.DOTALL,
                )
                cm = cell_pat.search(row_body)
                if not cm:
                    raise RuntimeError(f'cell {coord} not found in row {row_num}')

                style_m = re.search(r's="([^"]+)"', cm.group('a') or '')
                style = style_m.group(1) if style_m else None

                if kind == 'str':
                    new_cell = make_inline_cell(coord, value, style)
                elif kind == 'num':
                    new_cell = make_num_cell(coord, value, style)
                elif kind == 'bool':
                    new_cell = make_bool_cell(coord, value, style)
                else:
                    raise ValueError(kind)

                row_body = row_body[: cm.start()] + new_cell + row_body[cm.end():]
            rows_touched.add(row_num)

        return f'{row_open}{new_attrs}>{row_body}{row_close}'

    # Match each row preserving open tag, attrs, body, close tag
    new_xml = re.sub(
        r'(<row)([^>]*)>(.*?)(</row>)',
        patch_row,
        xml,
        flags=re.DOTALL,
    )

    # Sanity: every update must have landed
    expected_rows = {int(re.match(r'^[A-Z]+(\d+)$', c).group(1)) for c, _, _ in UPDATES}
    missing = expected_rows - rows_touched
    if missing:
        raise RuntimeError(f'rows not touched: {sorted(missing)}')
    return new_xml


def main() -> None:
    SRC_ABS = (Path(__file__).resolve().parent.parent / SRC).resolve()
    DST_ABS = (Path(__file__).resolve().parent.parent / DST).resolve()

    # Read input
    with zipfile.ZipFile(SRC_ABS, 'r') as zin:
        names = zin.namelist()
        contents: dict[str, bytes] = {n: zin.read(n) for n in names}

    # Patch sheet2.xml (DATOS GENERALES)
    sheet_xml = contents['xl/worksheets/sheet2.xml'].decode('utf-8')
    new_sheet = update_sheet(sheet_xml)
    contents['xl/worksheets/sheet2.xml'] = new_sheet.encode('utf-8')

    # Write atomically next to destination
    tmp = DST_ABS.with_suffix(DST_ABS.suffix + '.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    shutil.move(tmp, DST_ABS)
    print(f'OK · {DST_ABS}')


if __name__ == '__main__':
    main()
