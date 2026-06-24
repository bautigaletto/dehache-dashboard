# DEHACHE · Panel de Ventas 2026

Dashboard interactivo de ventas, alojado en GitHub Pages y actualizable subiendo el Excel.

## 🔗 Ver el dashboard

Una vez configurado, el link público será:
```
https://TU_USUARIO.github.io/TU_REPOSITORIO/
```

---

## ⚙️ Configuración inicial (solo la primera vez)

### 1. Crear repositorio en GitHub
1. En [github.com](https://github.com) → **New repository**
2. Nombre sugerido: `dehache-dashboard`
3. Dejarlo **Public** (necesario para GitHub Pages gratuito)
4. **No** inicializar con README

### 2. Subir estos archivos
Subí todos los archivos de esta carpeta al repositorio:

```
dehache-dashboard/
├── POWER_BI.xlsx           ← el Excel de datos
├── LogoDHVertical.jpg      ← logo de la empresa
├── template.html           ← plantilla del dashboard
├── update.py               ← script procesador
├── index.html              ← dashboard generado (se actualiza automáticamente)
└── .github/
    └── workflows/
        └── update.yml      ← automatización
```

Podés arrastrar todos los archivos desde la web de GitHub.

### 3. Habilitar GitHub Pages
1. En tu repositorio → **Settings** → **Pages**
2. En *Source* seleccioná **"GitHub Actions"**
3. Guardá

### 4. Habilitar permisos del workflow
1. Settings → **Actions** → **General**
2. En *Workflow permissions* → seleccioná **"Read and write permissions"**
3. Guardá

---

## 🔄 Actualizar los datos (uso diario)

### Opción A – Desde la web de GitHub (más simple)
1. Abrí tu repositorio en github.com
2. Hacé clic en `POWER_BI.xlsx` → **botón lápiz ✏️ o los tres puntos** → *Upload file*
3. Seleccioná el nuevo Excel → **Commit changes**
4. El dashboard se actualiza automáticamente en ~2 minutos

### Opción B – Desde tu computadora con git
```bash
git add POWER_BI.xlsx
git commit -m "Actualización datos $(date +%d/%m/%Y)"
git push
```

### Opción C – Actualizar manualmente sin subir el Excel
1. En tu repositorio → **Actions** → *Actualizar Dashboard*
2. **Run workflow** → Run

---

## 🛠️ Actualizar localmente (sin GitHub)

Si querés regenerar el dashboard en tu PC:

```bash
# Instalar dependencias (solo una vez)
pip install pandas openpyxl Pillow

# Reemplazá POWER_BI.xlsx con la versión nueva y ejecutá:
python update.py

# Abrí index.html en tu navegador
```

---

## 📋 Reglas de negocio aplicadas

| Exclusión | Motivo |
|-----------|--------|
| Nivel `Descuentos` | No es venta real |
| Nivel `Varios` | No se incluye en facturación |
| Vendedor `SUCURSAL CENTRO` | Interno |
| Centro `0005` | Interno |
| Artículos nivel `Oportunidades` | Excluidos del alerta de stock sin movimiento |

**Clientes sin compra reciente:** 3+ meses consecutivos sin comprar  
Excluye zonas: Clientes Especiales (17), Transporte DH (25), Empleados (96), Sucursal Centro (97), zona 15 y zona 21.

**Artículos sin movimiento:** más de 12 meses sin venta (historial desde marzo 2025).

---

## 📊 Objetivo mensual de facturación

| Nivel | Objetivo mensual |
|-------|-----------------|
| Impermeabilización | $110.000.000 |
| Materiales construcción | $69.000.000 |
| Sanitarios | $23.900.000 |
| Instalación agua y cloaca | $22.000.000 |
| Construcción en seco | $19.200.000 |
| Fletes | $14.000.000 |
| Herramientas | $7.800.000 |
| Aislantes | $7.000.000 |
| Pinturería | $6.600.000 |
| Químicos | $4.750.000 |
| Zinguería | $4.500.000 |
| Instalación de gas | $2.350.000 |
| Jardinería y riego | $250.000 |
| Elementos de seguridad | $175.000 |
| **TOTAL** | **$300.000.000** |
