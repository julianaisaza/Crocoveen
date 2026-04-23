# Crocoveen — Contexto del Proyecto

## ¿Qué es esto?
**Crocoveen** es el dashboard de gestión de proyectos inmobiliarios de **Tierra Grata & Co.** Es una app web de un solo archivo HTML (`crocoveen.html`) publicada en GitHub Pages. Se actualiza semanalmente con datos de Excel mediante un script Python (`sync_crocoveen.py`).

**URL pública:** https://julianaisaza.github.io/Crocoveen/crocoveen.html  
**Repositorio:** https://github.com/julianaisaza/Crocoveen.git  
**Carpeta local (Windows):** `C:\Users\JULIANA ISAZA\OneDrive - TIERRA GRATA & CO. S.A.S\Escritorio\Tierra Grata Comercial\Crocoveen\Crocoveen App\`

---

## Archivos clave

| Archivo | Descripción |
|---|---|
| `crocoveen.html` | App completa (HTML + CSS + JS en un solo archivo) |
| `sync_crocoveen.py` | Script Python que lee los Excel y actualiza el HTML |
| `Sync Crocoveen.bat` | Doble clic → sincroniza + git push + envía correo |
| `Sync Crocoveen (sin correo).bat` | Igual pero sin enviar el correo (para pruebas) |
| `config.json` | Credenciales de email (email_from, email_password) |
| `Actividades Proyectos.xlsx` | Define los hitos automáticos (está en la carpeta padre) |

**Carpeta padre** (`../`): contiene todos los Excel de Crocoveen por proyecto (ej. `TG Ambarte Crocoveen.xlsx`, `TG Summit Grand Crocoveen.xlsx`, etc.)

---

## Flujo de trabajo semanal

```
Abrir Sync Crocoveen.bat (Windows)
  → parse_file() lee cada *Crocoveen*.xlsx de la carpeta padre
  → update_html() inyecta los datos en crocoveen.html
  → push_to_github() hace git add + commit + push
  → send_reminder() envía correo HTML a todo el equipo
```

El `git push` publica en GitHub Pages (~1 min de delay). El script **nunca** reescribe el HTML desde cero — solo actualiza bloques específicos via regex/brace-counting.

---

## Arquitectura del HTML

`crocoveen.html` es una SPA con estado en el objeto `S`:

```javascript
const S = {
  view: 'dashboard',       // pestaña activa
  towers: [...],           // array de proyectos/torres
  rhythm: {...},           // ritmo de ventas por torre y mes
  deliveries: {...},       // entregas clientes por torre y mes
  constr: {...},           // fechas de construcción por torre
  milestones: [...],       // hitos automáticos
  currentDate: '2026-04',  // mes de referencia (auto-actualizado por sync)
  projFilter: [],          // filtro activo en Dashboard
  crocoProjFilter: [],     // filtro activo en Crocoveen
}
```

**Vistas:** Dashboard · Crocoveen (timeline) · Procesos · Tareas  
**Persistencia:** Supabase REST API (task_owners, task_completions, team_people)  
**Auth:** Hash SHA-256 de contraseña en `PW_HASH`

---

## Proyectos y torres

Cada torre tiene un `id` estable. El `TOWER_MAP` en `sync_crocoveen.py` mapea `(keyword_archivo, keyword_torre) → id`:

| Proyecto | Torres | IDs |
|---|---|---|
| Country Life | T1A, T1B, T2A, T2B | C-T1A, C-T1B, C-T2A, C-T2B |
| Verano | T1A, T1B, T2A, T2B | V-T1A, V-T1B, V-T2A, V-T2B |
| Zitizen | T1, T2 | Z-T1, Z-T2 |
| TG Mágica | T2A, T2B, T3A, T3B, T4A, T4B, T4C | M-T2A … M-T4C |
| Bosketo | T1, T2A, T2B, T3 | B-T1, B-T2A, B-T2B, B-T3 |
| Camino Verde | T3, T4 | CV-T3, CV-T4 |
| Ambarte | T1A, T1B, T1C | A-T1A, A-T1B, A-T1C |
| La Vida es Bella | T1–T6 | L-T1 … L-T6 |
| Summit | T1, T2 | S-T1, S-T2 |
| Summit Grand | T1, T2, T3 | SG-T1, SG-T2, SG-T3 |
| Primavera | T1, T2, T3 | P-T1, P-T2, P-T3 |
| Lúmina | T1, T2, T3 | Lu-T1, Lu-T2, Lu-T3 |

**norm(s):** `re.sub(r'[^a-z0-9]', '', str(s).lower())` — se usa para comparar keywords sin tildes/espacios/guiones.

---

## Estructura de los Excel de Crocoveen

Cada Excel tiene columnas trimestrales (ej. `Ene-Mar 26`, `Abr-Jun-26`) con función `build_col_month()` que las expande a meses individuales.

Por cada torre, las filas son:
1. **Torre** — `Unidades`, `Vendidas`, `Disponibles` + ritmo de ventas mensual (celdas verdes)
2. **% acumulado** — porcentaje de ventas
3. **C/E** — marcadores de construcción (`C`) y lista-entrega (`E`)
4. **Entregas clientes** — fila con texto que contiene "Cron. entregas" o "Crédito clientes" + valores numéricos por mes

**Detección de entregas:** La función escanea el texto de TODAS las celdas string de la fila:
```python
row_text = ' '.join(str(v).strip() for v in nrow if v is not None and isinstance(v, str)).lower()
if ('cron' in row_text or 'entregas' in row_text or 'crédito' in row_text or 'credito' in row_text):
```

---

## Hitos automáticos (Actividades Proyectos.xlsx)

El archivo tiene dos secciones:
- **POR PROYECTO** — un hito por proyecto completo (ej. Estudio de Títulos, Contrato Lote)
- **POR ETAPA** — un hito por torre (ej. apertura fiducia, inicio construcción)

Cada actividad tiene: `Nombre`, `Responsable`, `Periodo` (ej. "5 meses antes de lanzamiento", "1 mes antes de PE", "5 meses antes de entregas").

Los IDs de hitos son **determinísticos** via MD5:
```python
int(md5(f"{tid}_{tipo}").hexdigest(), 16) % 900000 + 100000
# Para hitos POR PROYECTO: f"~{pid}_{tipo}"
```

**Regla importante:** Los hitos con fecha pasada (< currentDate) no se convierten en tareas (se asumen completados).

**Lanzamiento:** `lanzamiento = min(rhythm.keys()) if rhythm and sold == 0 else None` — solo proyectos con 0 ventas tienen fecha de lanzamiento.

---

## Supabase

- **URL:** `https://vsxqlyhrakxqiwtdevsj.supabase.co`
- **Tablas:** `task_owners` (quién es responsable de cada tarea), `task_completions` (tareas marcadas como hechas), `team_people` (lista del equipo)
- **Acceso:** anon key embebida en el script y en el HTML

---

## Correo semanal

- Se envía via SMTP Gmail (configurado en `config.json`)
- Lista: jisaza, andresmesab, vroldan, jbarbosa, jsuarez, jarango, nparra, chincapie, aestrada, dsierra @tierragrata.co
- Colores corporativos: Banner `#0D3D52`, verde `#007060`
- Para omitir el correo: usar `Sync Crocoveen (sin correo).bat` o pasar `--no-email` al script

---

## Convenciones de código importantes

### `push_to_github(folder)`
Solo hace `git add crocoveen.html` (no el script Python). Si no hay cambios en el HTML dice "Sin cambios nuevos" y no pushea. Para pushear cambios del script Python, hacer `git add` manualmente desde la terminal.

### `update_html(all_results)`
Modifica bloques específicos del HTML usando `replace_js_block()` (brace-counting) y regex. **No reescribe el archivo completo** — preserva todo el JS/CSS personalizado.

### Colores del proyecto (PC)
```javascript
const PC = {
  'Country Life':'#2e7d32','Verano':'#558b2f','Bosketo':'#e65100',
  'Zitizen':'#1565c0','Summit':'#6a1b9a','Summit Grand':'#b71c1c',
  'Ambarte':'#00695c','La Vida es Bella':'#f57f17','Camino Verde':'#33691e',
  'Primavera':'#ad1457','TG Mágica':'#4527a0','Lúmina':'#0277bd'
}
```

---

## Historial de cambios relevantes

- **Detección entregas:** Escanea row_text de todas las celdas string (no solo las primeras 3 columnas)
- **Lanzamiento falso:** Corregido — solo proyectos con `sold == 0` reciben hitos de lanzamiento
- **currentDate:** Auto-actualizado al mes actual en cada sync
- **Barra de selección:** Al arrastrar sobre celdas del timeline aparece barra azul con Promedio/Recuento/Suma (como Excel)
- **POR PROYECTO vs POR ETAPA:** Actividades Proyectos.xlsx soporta ambos scopes
- **--no-email flag:** `python sync_crocoveen.py --no-email` para sincronizar sin enviar correo

---

## Cómo agregar un proyecto nuevo

1. Agregar entradas en `TOWER_MAP` con el keyword del archivo y de cada torre
2. Agregar el keyword en el loop de `file_keys` (línea ~149)
3. Agregar la torre en el array `towers` del HTML con `{id, proj, name, tot, sold}`
4. Agregar el color del proyecto en `PC` del HTML
5. Crear el Excel `*NombreProyecto* Crocoveen.xlsx` en la carpeta padre con la estructura estándar
6. Correr el sync

---

## Diagnóstico rápido

| Síntoma | Causa probable |
|---|---|
| Cambios no visibles en el sitio | Caché del navegador → Ctrl+Shift+R |
| "Sin cambios nuevos" en el bat | HTML no cambió; hacer git push manual si hay cambios del script |
| Entregas no aparecen | Verificar que la fila tenga texto con "cron", "entregas" o "crédito" |
| Hito en proyecto ya vendido | `sold > 0` → lanzamiento = None → no genera hitos de lanzamiento |
| Push desde carpeta equivocada | Siempre correr git desde `Crocoveen App\`, no desde `Crocoveen\` |
