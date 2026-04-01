# Instructivo de Formato — Reportes Comparativos BuyFlow (Excel)

## Objetivo
Este documento define las reglas exactas de formato, estructura y contenido para generar reportes comparativos de proveedores en Excel bajo la marca BuyFlow. El objetivo es que cada reporte salga lo más cercano posible al estándar final sin necesidad de ajustes manuales.

---

## 1. Estructura general del archivo

- **Una hoja por requerimiento** (o grupo de requerimientos cuando son el mismo producto/línea).
- **Nombre de la hoja**: descriptivo y corto. Ej: `Portobello Gouache 7,5x15,5`, `Revestimiento Negro 30 m²`.
- **Columnas**: B para etiquetas, C en adelante para proveedores (un proveedor por columna).
- **Columna A**: no se usa (queda como margen izquierdo).
- **Anchos de columna**: B = 35, columnas de proveedores = 28–30.

---

## 2. Alineación (CRÍTICO)

**Todas las celdas del reporte llevan alineación horizontal CENTER y vertical CENTER**, salvo que el contenido sea un párrafo largo (Requerimiento, Recomendación, Notas) en cuyo caso se agrega **wrap_text=True** y se mantiene center/center.

Esto aplica a:
- Etiquetas de la columna B
- Valores de datos en columnas de proveedores
- Headers de proveedor
- Fila de recomendación
- Notas

**Nunca dejar alineación por defecto (left/None).** Siempre setear explícitamente `horizontal='center', vertical='center'`.

---

## 3. Bordes (CRÍTICO)

El reporte usa **bordes medium** para separar visualmente las secciones. Las reglas son:

| Sección | Borde |
|---------|-------|
| Requerimiento (fila 3) | `border_top=medium` + `border_bottom=medium` en toda la fila con contenido |
| Recomendación BuyFlow (fila 5) | `border_top=medium` + `border_bottom=medium` en toda la fila con contenido |
| Proveedores analizados (fila 7) | `border_top=medium` + `border_bottom=medium` en toda la fila con contenido |
| Headers Proveedor A/B/C (fila 9) | `border_top=medium` |
| Modelo (fila 10) | `border_top=medium` (cierra la separación con header) |
| Costo total (fila 14) | `border_bottom=medium` (separa del bloque de ahorro) |
| Ahorro en $$$ (fila 15) | `border_top=medium` |
| Ahorro % (fila 16) | `border_bottom=medium` |
| CAPACIDAD Y LOGÍSTICA (fila 17) | `border_top=medium` |
| Plazo de entrega (fila 19) | `border_bottom=medium` |
| CONDICIONES COMERCIALES (fila 20) | `border_top=medium` |
| Financiación (fila 21) | `border_bottom=medium` |
| SERVICIO AL CLIENTE (fila 22) | `border_top=medium` |
| Garantía (fila 23) | `border_bottom=medium` |
| Proveedor Recomendado (fila 24) | `border_top=medium` + `border_bottom=medium` |
| Notas (fila 25) | `border_top=medium` + `border_bottom=medium` |

**Los bordes se aplican a TODAS las celdas de la fila que tengan contenido** (no solo a la columna B). Si la fila tiene datos en B, C, D, E — las 4 celdas llevan el borde.

---

## 4. Colores y fills

- **Naranja BuyFlow**: `FFB24E` (PatternFill solid). Se usa en:
  - B3 ("Requerimiento:")
  - B5 ("Recomendación BuyFlow")
  - B7 ("Proveedores analizados")
  - Fila 9 (headers "Proveedor A", "Proveedor B", etc.)
  - Fila 24 completa ("Proveedor Recomendado BuyFlow" + la celda con "X" o "Alternativa")
- **Texto en celdas naranjas**: usar color de font que contraste (theme color o blanco según el caso). En la práctica, el theme=2 funciona bien.
- **Resto de celdas**: sin fill (fondo transparente).

---

## 5. Tipografía

| Elemento | Font | Tamaño | Bold |
|----------|------|--------|------|
| Requerimiento (B3 y C3) | Arial | 14 | No |
| Recomendación BuyFlow (B5 y C5) | Arial | 14 | No |
| Proveedores analizados (B7) | Arial | 10 | No |
| Headers proveedor (fila 9) | Arial | 10 | No |
| Etiquetas columna B (filas 10–25) | Arial | 10 | Ver nota |
| Valores de proveedores | Arial | 10 | No |
| Nombre del proveedor (fila 12) | Arial | 10 | **Sí** |
| Costo total (fila 14, etiqueta y valores) | Arial | 10 | **Sí** |
| Ahorro $$$ y Ahorro % (etiquetas) | Arial | 10 | **Sí** |
| Títulos de sección (CAPACIDAD, CONDICIONES, SERVICIO) | Arial | 10 | **Sí** |
| Provincia (etiqueta) | Arial | 10 | **Sí** |

**Nota sobre bold en etiquetas B**: las etiquetas de sección (CAPACIDAD Y LOGÍSTICA, CONDICIONES COMERCIALES, SERVICIO AL CLIENTE Y SOPORTE) y las etiquetas de datos clave (Proveedor, Costo total, Ahorro, Provincia) van en bold. Las etiquetas informativas (Modelo, Precio por m², Envío, Plazo de entrega, Financiación, Garantía) van sin bold.

---

## 6. Formato de números

| Tipo | Formato Excel |
|------|---------------|
| Moneda (precios, costos, ahorros) | `_("$ "* #,##0_);_("$ "* \(#,##0\);_("$ "* \-??_);_(@_)` |
| Porcentaje (ahorro %) | `0%` |
| General (textos, "A definir", "-") | `General` |

**Importante**: usar el formato de moneda con "$ " (peso argentino con espacio) y sin decimales.

---

## 7. Fórmulas

Las fórmulas deben ser dinámicas y referenciar celdas, nunca hardcodear resultados.

| Campo | Fórmula tipo |
|-------|-------------|
| Costo total | `=+C11*42.95` (precio/m² × cantidad de m²) |
| Ahorro en $$$ | `=+C14-D14` (siempre contra el proveedor de referencia / columna C) |
| Ahorro % | `=1-(D14/C14)` (siempre contra el proveedor de referencia / columna C) |

**Reglas de ahorro**:
- La columna del proveedor de referencia (generalmente C) lleva "-" en ahorro $$$ y ahorro %.
- Si un proveedor no tiene precio definido ("A definir"), poner "-" en sus celdas de ahorro (no fórmula).
- Las fórmulas de ahorro siempre comparan contra la columna C ($C$14), no contra el proveedor anterior.

---

## 8. Heights de fila

| Fila | Height | Motivo |
|------|--------|--------|
| Filas vacías separadoras (2, 4, 6, 8) | 14 | Espaciado visual mínimo |
| Requerimiento (3) | 80–125 según largo del texto | Wrap text |
| Recomendación (5) | 100–140 según largo del texto | Wrap text |
| Modelo (10) | 28 | Doble línea si el nombre es largo |
| Envío (13) | 28 si hay texto largo, 14 si es corto | Adaptable |
| Notas (25) | 70–95 según contenido | Wrap text |
| Resto de filas | 14 (default) | |

---

## 9. Merged cells

Solo se usan merges en:
- **C3:E3** (o hasta la última columna de proveedor) — texto del requerimiento.
- **C5:E5** (o hasta la última columna de proveedor) — texto de la recomendación.

No se mergean más celdas.

---

## 10. Fila 24 — Proveedor Recomendado BuyFlow

- Toda la fila (B24 + celdas de proveedores con contenido) lleva **fill naranja FFB24E**.
- En la columna del proveedor recomendado se pone **"X"**.
- Si hay un proveedor que es alternativa válida (no el recomendado principal), se puede poner **"Alternativa"** en su celda, también con fill naranja.
- Las celdas de proveedores no recomendados quedan vacías (pero con fill naranja si están en el rango visual).

---

## 11. Contenido de la Recomendación BuyFlow (fila 5)

La recomendación debe:
- Ser concisa (máximo 5 líneas).
- Indicar claramente qué proveedor se recomienda y por qué.
- Mencionar el ahorro si aplica (en % y/o $$$).
- Si el producto recomendado es el exacto solicitado, decirlo explícitamente: "Se recomienda proveedor X ya que es el único que cumple con el requerimiento exacto".
- Si hay alternativas más baratas pero en otro formato/modelo, mencionarlas como alternativa y dejar la decisión al cliente.
- No inventar datos. Si algo no se sabe, no se incluye.

---

## 12. Proveedores analizados (C7)

- El número en C7 refleja la **cantidad total de proveedores contactados/analizados** durante la búsqueda, no solo los que aparecen en la tabla.
- Si se contactaron 9 proveedores pero solo 4 tienen datos relevantes para mostrar, C7 = 9 y la tabla muestra 4 columnas.

---

## 13. Notas y Observaciones (fila 25)

- Cada proveedor puede tener su nota en su columna correspondiente.
- Las notas deben ser factuales y útiles para la decisión: formato alternativo, disponibilidad de stock, condiciones especiales, limitaciones del producto.
- Si un proveedor ofrece un producto que no es el exacto solicitado, la nota debe decir "Formato alternativo (detalle). No es el modelo exacto solicitado."
- Si hay algo negociable o una acción posible, mencionarlo: "En caso de analizarlo como opción, se puede negociar el precio."

---

## 14. Checklist rápido antes de entregar

- [ ] Todas las celdas con contenido tienen alineación center/center
- [ ] Bordes medium aplicados en todas las filas separadoras (y en TODAS las celdas de esas filas)
- [ ] Fill naranja FFB24E en B3, B5, B7, fila 9, fila 24
- [ ] Fórmulas dinámicas (no valores hardcodeados) en costo total, ahorro $$$ y ahorro %
- [ ] Formato de moneda `_("$ "* #,##0...)` en todas las celdas de precios/costos/ahorros
- [ ] Formato `0%` en celdas de ahorro %
- [ ] Wrap text activado en requerimiento, recomendación, modelo, envío (si es largo), y notas
- [ ] Heights de fila ajustados para que el contenido sea legible
- [ ] Nombre de hoja descriptivo
- [ ] Número de proveedores analizados correcto en C7
- [ ] Recomendación BuyFlow completa y concisa (≤5 líneas)
- [ ] Recalcular fórmulas con `scripts/recalc.py` antes de entregar
