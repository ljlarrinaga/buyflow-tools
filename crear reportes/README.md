# Crear Reportes BuyFlow

Genera reportes comparativos de proveedores en formato `.xlsx` con el estilo visual de BuyFlow.

---

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `buyflow.py` | Script generador y validador del reporte Excel |
| `Instructivo_Formato_Reportes_BuyFlow.md` | Especificaciones de formato (colores, bordes, tipografía) |
| `Reglas_Negocio_BuyFlow.md` | Reglas de negocio (ahorro, referencia, selección de proveedores) |
| `Logo Buyflow.png` | Logo que se inserta automáticamente en el reporte |

---

## Requisitos

```
pip install openpyxl
```

---

## Tipos de reporte

**Reporte estándar** — comparación con precio unitario. Campos: modelo, precio, nombre, envío, provincia, plazo, financiación, garantía, notas.

**Con descripción** (`tiene_descripcion: True`) — agrega fila de specs técnicas bajo el modelo. Usar cuando el nombre del modelo no es suficiente.

**Con filas custom** (`custom_rows`) — para atributos técnicos específicos (diámetro, potencia, peso, etc.).

**Por m²** — se configura `costo_formula` para calcular costo total según cantidad.

---

## Columna de Referencia (col C)

- Si el cliente tiene proveedor actual → ese va como Referencia en col C
- Si no tiene → Claude busca el precio promedio de mercado y lo usa como Referencia
- Siempre hay una referencia; el ahorro siempre se calcula contra col C

---

## Cómo usar con Claude Code

Pasale a Claude: el requerimiento, los datos de proveedores, si hay referencia del cliente, y la cantidad.

Claude va a leer los docs, armar el config, generar el `.xlsx` y validarlo automáticamente.

**Ejemplo de prompt:**
```
Consultá los docs en "crear reportes" de buyflow-tools y generame un reporte:

Requerimiento: cerámica 30x30 beige para 42 m²
Referencia del cliente: Porcelanosa $8.500/m² c/IVA

Proveedores:
- Easy: $6.200/m² c/IVA, envío $3.000, stock inmediato
- Sodimac: $6.800/m² c/IVA, envío gratis, 3 días
```

---

## Reglas clave

- Mostrar 3-4 proveedores más significativos en la tabla (el resto cuenta para "proveedores analizados")
- `proveedores_analizados` default: 8
- Todos los precios en la misma base (c/IVA o s/IVA)
- Links: incluir siempre que estén disponibles (`tiene_link: True`)
