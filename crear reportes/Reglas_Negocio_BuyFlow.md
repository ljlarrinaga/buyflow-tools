# Reglas de Negocio BuyFlow — Para Claude

## Cómo usar este proyecto

### Archivos en la carpeta del proyecto:
1. **buyflow.py** — Generador + Validador de reportes Excel
2. **Instructivo_Formato_Reportes_BuyFlow.md** — Especificaciones de formato detalladas
3. **Copia_de_LogoBF_fondo_blanco.png** — Logo (se inserta automáticamente)
4. **Reglas_Negocio_BuyFlow.md** — Este archivo

### Flujo de trabajo (para Claude):
1. El usuario pasa requerimiento + datos de proveedores
2. Claude arma el dict `config` según la estructura de buyflow.py
3. Claude ejecuta `generate_report(config)` 
4. Claude ejecuta `validate_file()` automáticamente
5. Si PASS → entrega el .xlsx
6. Si FAIL → corrige y re-valida sin preguntar

### NO hacer:
- No releer el instructivo ni buyflow.py cada vez (ya está internalizado)
- No reescribir código de formato (usar la función)
- No explicar el código paso a paso (ser conciso)

---

## Reglas de negocio

### Proveedores analizados
- Siempre poner **8** en la celda C7, salvo que el usuario indique otro número explícitamente.

### Referencia
- Si el cliente tiene proveedor actual → ese va como Referencia (col C).
- Si NO tiene → buscar el precio promedio de mercado del producto (1 búsqueda web) y usarlo como referencia con nombre descriptivo.
- La referencia SIEMPRE es algo real (proveedor o precio de mercado), nunca un cálculo derivado.

### Ahorro
- Siempre mostrar ahorro $ y % contra la Referencia (col C).
- Si no hay referencia del cliente, el ahorro se calcula contra el precio de mercado.

### Selección de proveedores a mostrar
- Si el usuario pasa muchos proveedores (ej: 8), seleccionar los **3-4 más significativos** para la tabla.
- NO incluir un proveedor carísimo sin justificación — solo si tiene algún valor diferencial (ej: única opción que cumple specs exactas, garantía especial, etc.).
- Los proveedores descartados cuentan para el número de "proveedores analizados" pero no aparecen en la tabla.

### Links
- Si tenemos links (de ML u otras plataformas), incluir fila Link siempre — tanto en reportes FCD como normales.
- Configurar `"tiene_link": True` en el config.

### Descripción
- Solo incluir fila Descripción cuando hay info técnica relevante que no cabe en Modelo (specs detalladas, materiales, medidas).
- Si el modelo ya es suficientemente descriptivo, no agregar Descripción innecesaria.

### Filas custom
- Cuando el producto tiene atributos técnicos específicos relevantes para la comparación (Diámetro, Potencia, Peso, Medidas, etc.), agregarlos como `custom_rows`.
- Se insertan entre Provincia y Plazo de entrega.

### Recomendación
- Máximo 5 líneas.
- Priorizar cumplimiento técnico sobre precio.
- Mencionar ahorro % si aplica.
- No inventar datos.

### IVA
- Normalizar todos los precios a la misma base (s/IVA o c/IVA).
- Indicar en `precio_label` cuál se usó.
- Si hay mezcla, convertir y anotar en Notas.

---

## Estructura del config dict

```python
config = {
    # Layout
    "tiene_descripcion": False,     # True si hay info técnica detallada
    "tiene_link": True,             # True si hay links de productos
    "custom_rows": [],              # Filas extra opcionales
    
    # Contenido
    "sheet_name": "Nombre corto",
    "requerimiento": "Se solicita...",
    "recomendacion": "Se recomienda...",
    "proveedores_analizados": 8,
    "cantidad": 1,
    "precio_label": "Precio c/iva",
    "costo_total_label": "Costo Total",
    "costo_formula": "=+{col}{price_row}*{qty}",
    
    # Heights (opcionales, tienen defaults razonables)
    "req_height": 120,
    "rec_height": 140,
    "modelo_height": 42,
    "desc_height": 60,
    "notas_height": 80,
    
    # Proveedores (el primero siempre es Referencia)
    "proveedores": [
        {
            "header": "Referencia",
            "modelo": "...",
            "descripcion": "...",
            "link": "https://...",
            "precio": 50000,
            "nombre_proveedor": "...",
            "envio": 0,
            "provincia": "...",
            "plazo_entrega": "...",
            "financiacion": "...",
            "garantia": "...",
            "recomendado": None,  # "X", "Alternativa", o None
            "notas": None,
            "custom": [],  # Valores para custom_rows
        },
    ]
}
```
