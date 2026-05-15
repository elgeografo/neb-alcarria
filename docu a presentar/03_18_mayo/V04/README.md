## V04 · base ALMUNIA 4.0 · estado y pendientes

**Fecha:** 2026-05-14. **Origen:** copia quirúrgica de V03 con las correcciones acordadas.
**Convocatoria:** RD 251/2024 · Intervención supraautonómica 7201 · Programa I.B · Cierre 18 mayo 2026 14:00 h

### Ficheros incluidos en V04 (renombrados ALMUNIA 4.0, sin PDFs)

- `Anexo_I_SOLICITUD_PRESUP_Form.xlsm` — solicitud + presupuesto (Anexo I oficial)
- `Anexo_II_B_PROPUESTA_ACTIV_PIB_ALMUNIA-4.0.docx` — propuesta de actividades (editado V04)
- `Anexo_III_C_CRONOGRAMA_ALMUNIA-4.0.xlsx`
- `Anexo_IV_D_DEC_RESPONSABLE_ALMUNIA-4.0.docx`
- `Anexo_V_E_DOC_VINCULANTE_ALMUNIA-4.0.docx`
- `Anexo_IX_J_DEC_COLABORACION_ALMUNIA-4.0.docx`
- `Anexo_XXII_MEMO_PRESUP_ALMUNIA-4.0.xlsx` — documento interno de soporte (no se sube)

---

## Instrucciones que se pidieron para el paso a V04

1. Renombrar ficheros: sufijo `NEB-ALCARRIA` → `ALMUNIA-4.0`. ✅ **Hecho** para todos los entregables.
2. No crear PDFs. ✅ **Hecho** (V04 contiene solo docx + xlsx + xlsm).
3. Completar el Anexo I (sólo 3 datos rellenos). ⏳ **Pendiente.**
4. Tabla del Anexo II (índice) desactualizada — páginas y epígrafes. ✅ **Resuelto.** El índice es un campo TOC automático de Word; en `word/settings.xml` se ha añadido `<w:updateFields w:val="true"/>` para que **Word actualice el TOC al abrir el documento** (también se puede forzar con click derecho → "Actualizar campo" o F9 sobre el índice).
5. En MEDIOS MATERIALES la sección "Justificación" hablaba de **familias** (concepto antiguo). ✅ **Resuelto:** sustituido por **Living Labs LL-3/LL-4/LL-5** (fila Parcela) y **LL-1/LL-2/LL-4** (fila Laboratorios UPM).
6. "Soporte para las **29 ediciones**" venía del modelo antiguo de 25 LLs. ✅ **Resuelto:** ahora dice **"10 ediciones (7 Living Labs · 45 días-actividad)"**.
7. Referencia en 2.7 a `NEB-ALCARRIA-PRESUPUESTO.xlsx` y "Anexo XXII Memoria del Presupuesto" (ambos internos). ✅ **Resuelto:** ambas referencias (Tabla 24 y cierre de fichas) ahora apuntan al **Anexo I — Solicitud y Presupuesto** (anexo oficial público).
8. En Justificación temporal aparecía "A.5 Coordinación P1" como si fuera una actividad. ✅ **Resuelto:** se ha reformulado en P1 y P2 como **"Imputación de la partida A.5 (coordinación transversal · no es actividad demostrativa)"** para que el evaluador no confunda partida con actividad.

---

## Cambios aplicados sobre V03 (resumen técnico)

- `word/settings.xml`: `<w:updateFields w:val="true"/>` (insertado en posición de schema válida, antes de `hdrShapeDefaults`).
- `word/document.xml`: 6 sustituciones de texto puntuales preservando runs y formato:
  - Tabla MEDIOS MATERIALES — fila Parcela (col Justificación)
  - Tabla MEDIOS MATERIALES — fila Laboratorios UPM (col Justificación)
  - Tabla MEDIOS MATERIALES — fila Residencia COPRODELI (col Justificación)
  - Tabla 2.7 Desglose por partida — primera línea
  - Tabla 2.7 Justificación temporal — bullet A.5 P1
  - Tabla 2.7 Justificación temporal — bullet A.5 P2
  - Párrafo de cierre de fichas — referencia final

Validación XML OK. 76 tablas y 983 párrafos preservados.

---

## Pendiente para V04

- **Anexo I** — completar todos los datos (sólo hay 3 rellenos según indicó el consorcio).
- (Opcional) Reconciliar las cifras del docx 2.7 con el último `Anexo_I_SOLICITUD_PRESUP_Form.xlsm` — algunos importes del cuerpo del Anexo II.B podrían no coincidir con la última actualización del Anexo I.
- (Si procede) Sustituir las menciones a "NEB-ALCARRIA" dentro del texto por "ALMUNIA 4.0" allá donde aún queden; los renombrados de fichero ya están hechos.
