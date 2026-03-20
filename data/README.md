# /data — Fuente de Datos PROFAR Dashboard

## Archivos de entrada

| Archivo | Fuente | Frecuencia | IVA |
|---------|--------|-----------|-----|
| `Resumen Mensual.xlsx` | Power BI PROFAR · E-COMMERCE PROFAR | Mensual | ❌ Sin IVA |
| `Detalle de Documentos.xlsx` | Power BI PROFAR · E-COMMERCE PROFAR | Mensual | ❌ Sin IVA |

## ¿Cómo actualizar los datos?

### Proceso mensual (manual → automático)

1. Ir a Power BI PROFAR
2. Filtros: `Estratificación = E-COMMERCE PROFAR` · `VENTA DE MEDICAMENTOS` · `Mes Completo`
3. Exportar **"Resumen Mensual"** → guardar como `Resumen Mensual.xlsx`
4. Exportar **"Detalle de Documentos"** → guardar como `Detalle de Documentos.xlsx`
5. Reemplazar los archivos en esta carpeta `/data/`
6. Hacer `git commit + push`
7. **El dashboard se actualiza solo en ~2 minutos** (GitHub Actions)

### Proceso automático (ya configurado)
- GitHub Actions corre el script 3 veces al día: **08:00 / 13:00 / 18:00 hora Chile**
- Workflow: `.github/workflows/update-dashboard.yml`
- Script: `scripts/update_dashboard.py`

## Reconciliación IVA

```
Power BI (sin IVA) × 1.19 = equivalente Magento (con IVA)

Ejemplo Feb 2026:
  Power BI: $92.4M sin IVA
  Magento:  $110.0M con IVA  (92.4 × 1.19 = 110.0 ✅ calza con Magento real)
```

## Estructura de datos Power BI esperada

### Resumen Mensual.xlsx
| Año | Mes | Rot. | Uni.Venta | Venta | Margen CR | % Margen CR | ... |
|-----|-----|------|-----------|-------|-----------|-------------|-----|
| 2025 | ENERO | 0.82 | 831 | 117,200,000 | 17,109,000 | 14.6% | ... |

### Detalle de Documentos.xlsx
| FECHA | CLIENTE | SKU | PRODUCTO | VENTA | MARGEN CR | NOMBRE_REGION | COMUNA | ... |
|-------|---------|-----|---------|-------|-----------|---------------|--------|-----|
| 31/01/2025 | Juan... | 110816 | NURTEC 75MG | 80,000 | ... | Región Metropolitana | Las Condes | ... |

## Logs de actualización

Cada actualización agrega un timestamp en los comentarios del HTML:
```javascript
// auto-update: 2026-03-20 18:00
```
