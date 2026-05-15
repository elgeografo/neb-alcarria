# V03 · Anexo II.B · estado y pendientes

**Documento:** `Anexo_II_B_PROPUESTA_ACTIV_PIB_NEB-ALCARRIA.docx` (plantilla oficial MAPA PIB rellenada con narrativa de 7 Living Labs + cuaderno COPRODELI-IA)
**PDF de vista previa:** `Anexo_II_B_PROPUESTA_ACTIV_PIB_NEB-ALCARRIA.pdf` (28 páginas A4)
**Convocatoria:** RD 251/2024 · Intervención supraautonómica 7201 · Programa I.B · Cierre 18 mayo 2026 14:00 h
**Fecha de actualización V03:** 2026-05-12

---

## Cambio narrativo principal (V02 → V03)

V03 incorpora la decisión estratégica de **cambio de relato** acordada con el consorcio:

- **Pieza protagonista**: **Cuaderno digital COPRODELI-IA** (aportación de IP de COPRODELI en TRL avanzado, con módulo de inteligencia artificial embebida para apoyo a la toma de decisión del agricultor). Sustituye al narrativo anterior de "construimos un cuaderno open-source desde cero".
- **Trayectoria TRL**: **TRL 7 → TRL 8 (cerca de TRL 9)**. Sustituye al TRL 5→7→8 anterior.
- **Estructura del proyecto**: **7 Living Labs (LL-0 .. LL-6)** sustituye al modelo anterior de 25 actividades en 10 familias.
- **B.1 Viajes**: modelo **persona-día** (923 pd manutención + 44 pd media dieta + 615 pernoctas + 27 vehículo-viajes flota COPRODELI). Sustituye al modelo "asistencia-edición".
- **Flota COPRODELI**: 4 vehículos con **20 plazas en TOTAL** (no 80). Corrección incorporada en todos los textos.
- **Presupuesto consolidado**: ≈ **215.711 €** (vs ≈ 185.350 € en V02). Distribución P1 ≈ 54.792 € (25 %) · P2 ≈ 160.919 € (75 %).

---

## Cómo se generó V03

V03 se construyó por **edición quirúrgica de la plantilla MAPA PIB** (la misma que V02), NO por regeneración pandoc:

1. **Plantilla base**: copia de V02 (que ya estaba sobre la plantilla oficial `02_templates/Anexo_II_B_PROPUESTA_ACTIV_PIB_(Nombre entidad_agrupación).docx`).
2. **Desempaquetado XML** de la docx vía `~/.claude/skills/docx/scripts/office/unpack.py`.
3. **Ediciones quirúrgicas en `word/document.xml`** — sustituciones de párrafos y celdas individuales preservando toda la maquetación, logos, tablas, checkboxes y estructura del template oficial MAPA.
4. **Re-empaquetado** vía `pack.py` con validación OK.
5. **PDF preview** vía LibreOffice (`soffice --headless --convert-to pdf`).

Resultado: documento idéntico a una plantilla MAPA rellenada manualmente en Word, sin artefactos de conversión.

---

## Los 7 Living Labs (sustituyen las 25 actividades de V02)

| Id | Living Lab | Duración | Edic. | Periodo · ventana |
|----|------------|---------|------|-------------------|
| **LL-0** | Concepto NEB + horticultura morisca alcarreña | 2 días | 1 | P1 · ene 2027 |
| **LL-1** | CUE+SIG · itinerario digital del hortelano (cuaderno **COPRODELI-IA** + integración CUE/SIGPAC) | 10 días (2 sem.) | 1 | P1 · feb–abr 2027 |
| **LL-2** | Drones agrarios · RGB (cartografía + regadío) + multiespectral (NDVI fitosanitario) | 4 días | 1 | P2 · sep 2027 |
| **LL-3** | Domo NEB · diseño digital + co-construcción | 10 días (2 sem.) | 1 | P2 · jul 2027 |
| **LL-4** | Living Lab digital · sensórica IoT + fotovoltaica + comunicaciones | 10 días (2 sem.) | 1 | P2 · ago–sep 2027 |
| **LL-5** | Ciclo productivo estacional (4 sesiones distribuidas) | 4 × 2 días | 4 | P2 · sep'27 · dic'27 · feb'28 · abr'28 |
| **LL-6** | Jornada final NEB-Alcarria · resultados, AKIS y replicabilidad | 1 día | 1 | P2 · abr 2028 |
| | **TOTAL** | **45 días-actividad** | **10 ediciones** | |

---

## 🟢 Lo que YA está bien en V03

### Portada y maquetación
- Logos oficiales (Cofinanciado UE + Gobierno España MAPA) intactos de la plantilla
- Encabezado "ANEXO II.B · PROPUESTA DE ACTIVIDADES — PROGRAMA I.B" preservado
- Tablas MAPA con bordes y formato oficial (T1, T5, T7 necesidades, T8 objetivos, T9 resultados, T12 listado actividades, T17 TIC, T22 presupuesto)
- Pie de página "Propuesta de actividades. Programa I.B" + numeración

### Contenido actualizado a la nueva narrativa

**§1 Resumen ejecutivo**
- "ACTIVIDADES PROPUESTAS": 7 Living Labs · 45 días-actividad · 10 ediciones
- "CONTENIDO": cuaderno COPRODELI-IA como pieza protagonista, TRL 7→8 (cerca de 9), 2 ciclos productivos
- "PROPUESTA DE VALOR DIFERENCIAL": 5 elementos reescritos con énfasis en COPRODELI-IA + 20 plazas en total

**§2.1 Experiencia previa**
- Bloque "Enfoque agrupación UPM" actualizado (integración COPRODELI-IA con CUE/SIGPAC)
- Bloques de proyectos previos siguen PENDIENTE (igual que V02, no se ha tocado)

**§2.2 Capacidad y estructura · Medios materiales**
- Fila reformulada como "Cuaderno digital COPRODELI-IA + stack auxiliar open-source UPM"
- TRL de partida actualizado a "Cuaderno COPRODELI-IA: TRL 7"

**§2.3 Necesidades · Objetivos · Resultados**
- **N5** reescrita: cuaderno digital con apoyo decisional por IA, no open-source genérico
- **O1** reescrita: validar y transferir el COPRODELI-IA
- **O6** reescrita: llevar el cuaderno COPRODELI-IA de TRL 7 a TRL 8
- **R1** reescrita: cuaderno COPRODELI-IA validado en Living Lab alcarreño
- **R2** reescrita: stack QGIS+QField preconfigurado, exporta a COPRODELI-IA
- **R7** reescrita: COPRODELI-IA en TRL 8 (cerca de TRL 9)
- **Justificación · Innovación**: integración COPRODELI-IA con ecosistema regulatorio español
- **Adopción por agricultores**: "7 Living Labs presenciales (45 días-actividad)"

**§2.4 Listado de actividades**
- Tabla MAPA T12 reducida de **25 filas (A1–A25) a 7 filas (LL-0..LL-6)**
- Todas las celdas (Id · Tipo · Nombre · Duración · Ediciones · N · O · R) rellenadas con el nuevo modelo
- **Usuarios finales**: 240 asistencias-edición, 80–110 participantes únicos, modelo persona-día explicado
- **Plan seguimiento**: referencias a actividades actualizadas a refs de LL

**§2.5 Empleo de TIC**
- Refs `A2/A3/A24` → `LL-1` (cuaderno COPRODELI-IA)
- Refs `A4/A25` → `LL-1, LL-2, LL-4`
- Refs `A4/A17` → `LL-2` (drones)
- Refs `A5/A6/A7/A16` → `LL-4` (IoT)
- Refs `A5–A22` → `LL-4, LL-5 (continuo)`
- Refs `A8/A9` → `LL-3` (cálculo estructural)
- Refs `A1/A23` → `LL-0, LL-6`
- Descripciones reescritas con COPRODELI-IA como pieza protagonista

**§2.6 Impacto · TRL story**
- "Modelo replicable y soberano" con 2 capas tecnológicas (COPRODELI-IA + stack auxiliar)
- TRL 5/7 antiguo → TRL 7/8 (cerca de 9) con referencia explícita a InfoDay 2026 ("no es I+D, es transferencia")
- Vector AKIS investigadores actualizado

**§2.7 Presupuesto**
- TOTAL: 185.350 € → **215.711 €**
- A.1 1.260 € (sin cambio; refs ahora LL-0 y LL-6)
- A.2 11.340 € (sin cambio; refs ahora LL-3 y LL-4)
- A.3 13.260 € → **15.445 €** (~184 h doc + ~353 h prep, 7 Living Labs)
- A.4 13.690 € → **28.047 €** (incluye horas de integración cuaderno COPRODELI-IA en LL-1)
- A.5 25.700 € → **26.525 €** (con tope 13.0 % directos)
- B.1 27.490 € → **63.719 €** (modelo persona-día completo)
- B.3 2.100 € → **2.800 €** (80 participantes únicos)
- C.1 6.880 € → **6.882 €** (refs a LLs actualizadas)
- C.2 36.000 € (sin cambio)
- Indirectos 9.500 € → **12.393 €** (15 % de personal directo recalculado)
- **Justificación temporal**: P1 nov 2026–may 2027 (~54.792 €) + P2 jul 2027–mar 2028 (~160.919 €) reescrita con LLs

---

## 🟡 Lo que SIGUE PENDIENTE para V04

### Sección 3 · Fichas de actividad
V02 sólo tenía la **ficha-modelo vacía** del template (T25–T41), sin desarrollar A1, A24, A10 ni el resto. V03 mantiene este estado. Para V04 hay que:
- Duplicar la ficha-modelo **7 veces** y rellenar cada Living Lab (LL-0 a LL-6)
- Cada ficha: descripción, tipo, tecnologías habilitadoras, ediciones (fecha/ubicación/duración), objetivos específicos, programa previsto, usuarios finales, recursos (RR.HH, materiales, técnicos), presupuesto detallado

### Datos del consorcio (igual que V02)

**🔴 COPRODELI (bloqueante):**
- Razón social oficial · CIF · domicilio · teléfono · email · web
- Ámbito territorial · exención de IVA
- Representante legal: apellidos, nombre, DNI, teléfono, email, cargo
- Persona de contacto técnico
- **Datos técnicos del cuaderno COPRODELI-IA (nuevo)**: nombre comercial, base tecnológica, descripción del módulo de IA, contextos previos de despliegue, métricas de validación previa, modelo de licencia/cesión

**🔴 UPM (bloqueante):**
- Razón social que firma · CIF · domicilio · teléfono · email · web
- Régimen de IVA · representante legal autorizado · IP del proyecto
- Resolución/poder de delegación

### Contenido a desarrollar
- §1.1 Descripción de entidades · 1 cara A4 por entidad (COPRODELI · UPM/GeoSo2)
- §2.1 Experiencia previa · 3–5 proyectos por entidad en digitalización y en sostenibilidad/género
- §2.2 Estructura organizativa, ETC, % mujeres, ubicación sede COPRODELI, plazas residencia
- §2.4 Plan promoción · cooperativas concretas, canales RRSS específicos, URL provisional de la web

---

## Cifras clave consolidadas V03

| Partida | Importe |
|---------|--------:|
| **TOTAL P1** (nov 2026 – may 2027) | ≈ **54.792 €** (25 %) |
| **TOTAL P2** (jul 2027 – mar 2028) | ≈ **160.919 €** (75 %) |
| **TOTAL GENERAL solicitado** | ≈ **215.711 €** |
| · A.1 Personal docente externo | 1.260 € |
| · A.2 Personal de apoyo externo | 11.340 € |
| · A.3 Personal docente propio | 15.445 € |
| · A.4 Personal de apoyo propio | 28.047 € |
| · A.5 Personal de coordinación propio | 26.525 € |
| · B.1 Viajes y dietas (persona-día) | 63.719 € |
| · B.2–B.7 Otros materiales | 14.100 € |
| · C.1 Inventariables | 6.882 € |
| · C.2 Otras inversiones | 36.000 € |
| · Indirectos (15 % personal directo) | 12.393 € |
| **Tope subvención** | 70.000 € ≤ X ≤ 500.000 € ✅ |
| **Topes de partida** | A.5 13 % · B.1 31 % · C 21 % — todos OK |

---

## Flujo de versionado

- V01 — batería completa de plantillas con pre-relleno básico (25 actividades)
- V02 — Anexo II.B sobre plantilla oficial PIB rellenado profundamente (25 actividades, ≈ 185 k €)
- **V03 (esta carpeta)** — Cambio de relato a 7 Living Labs + cuaderno COPRODELI-IA + TRL 7→8 + presupuesto persona-día (≈ 215 k €). Generado por edición quirúrgica del XML preservando la maquetación oficial MAPA al 100 %.
- V04 — desarrollar las 7 fichas LL (sec. 3) + insertar datos UPM/COPRODELI cuando lleguen
- V_final — última revisión consultor antes del 18 mayo 14:00 h
