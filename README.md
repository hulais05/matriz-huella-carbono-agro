# Matriz de Huella de Carbono — Sector Agropecuario

Sistema para relevar datos y calcular la huella de carbono corporativa de un establecimiento agropecuario (agricultura + ganadería), con un dashboard interactivo que se actualiza automáticamente a medida que se completa la planilla Excel.

🔗 **Demo en vivo:** _(pendiente de publicar — ver sección "Despliegue")_

## Despliegue (Streamlit Community Cloud)

Para compartir el dashboard con clientes mediante un enlace público y permanente:

1. Ingresar a [share.streamlit.io](https://share.streamlit.io) con la cuenta de GitHub.
2. "New app" → seleccionar el repo `hulais05/matriz-huella-carbono-agro`, branch `main`, archivo principal `dashboard_carbono.py`.
3. Deploy. Streamlit Cloud queda con una URL fija (ej. `https://matriz-huella-carbono-agro.streamlit.app`) que se actualiza sola con cada `git push`.
4. En la nube no existe la carpeta `~/Downloads`, así que el dashboard usa automáticamente la planilla de ejemplo incluida en el repo (`Matriz_HuellaCarbono_Agropecuario.xlsx`).

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `Matriz_HuellaCarbono_Agropecuario.xlsx` | Planilla de relevamiento de datos (10 hojas) |
| `dashboard_carbono.py` | Dashboard interactivo en Streamlit |
| `generar_matriz_carbono.py` | Script que genera la planilla Excel desde cero |
| `requirements.txt` | Dependencias de Python |
| `.streamlit/config.toml` | Configuración del tema del dashboard |

## Perfil del establecimiento

- Agricultura + ganadería, 1.000 ha totales / 900 ha productivas, un solo establecimiento
- Maquinaria: 2 tractores, 5 cosechadoras, 2 camiones (combustión a gasoil)
- Energía: 20% red eléctrica, 80% autogeneración a gasoil + zeppelin de GLP
- Agricultura: fertilizantes nitrogenados, urea, agroquímicos
- Ganadería: feedlot + tambo, con efluentes
- Alcance: Scope 1 + Scope 2, incluye sumideros (monte nativo + pasturas)
- Sin registros históricos de consumo ni software de gestión previo

## Metodología y normativa de referencia

- **GHG Protocol** (Corporate Standard) — enfoque de control operacional
- **ISO 14064-1:2018** — especificación para inventarios de GEI a nivel organizacional
- **IPCC 2006** (Vol. 2 Combustión, Vol. 4 Agricultura/Ganadería, Vol. 5 Efluentes) — factores de emisión Tier 1
- **IPCC AR6 (2021)** — Potenciales de Calentamiento Global a 100 años:
  - CH4 fósil (combustión gasoil/GLP): **29,8**
  - CH4 biogénico (fermentación entérica, estiércol, efluentes): **27,2**
  - N2O: **273**
- **CAMMESA** — factor de emisión de la red eléctrica argentina: 0,383 kg CO2e/kWh (2023). ⚠️ Verificar el valor vigente cada año en [cammesaweb.cammesa.com/download/factor-de-emision](https://cammesaweb.cammesa.com/download/factor-de-emision/)

## Categorías de emisión cubiertas

**Scope 1**
- S1-1 Combustión móvil (tractores, cosechadoras, camiones)
- S1-2 Combustión estacionaria (generadores + zeppelin GLP)
- S1-3 Ganadería (fermentación entérica + gestión de estiércol)
- S1-4 Suelos agrícolas (N2O por fertilizantes/urea/residuos)
- S1-5 Efluentes (feedlot + tambo)

**Scope 2**
- S2-1 Electricidad de red (mensual, 12 meses)

**Sumideros**
- Monte nativo
- Pasturas implantadas

## Cómo usar

### 1. Completar la planilla Excel
Abrir `Matriz_HuellaCarbono_Agropecuario.xlsx` y completar las celdas de color:
- 🟡 **Amarillo** → datos a ingresar
- 🟢 **Verde** → calculado automáticamente (no modificar)
- 🔵 **Celeste** → factor de referencia (no modificar)

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Ejecutar el dashboard
```bash
streamlit run dashboard_carbono.py
```

Por defecto el dashboard busca el Excel en `~/Downloads/Matriz_HuellaCarbono_Agropecuario.xlsx`. La ruta se puede cambiar desde la barra lateral.

### 4. Actualización en vivo
Con "Auto-actualizar al guardar Excel" activado, cada vez que se guarda la planilla (`Ctrl+S`) el dashboard detecta el cambio y recalcula todos los indicadores y gráficos sin necesidad de reiniciar.

## Regenerar la planilla desde cero

Si se necesita una planilla nueva (vacía) con la misma estructura:
```bash
python3 generar_matriz_carbono.py
```
Esto genera `Matriz_HuellaCarbono_Agropecuario.xlsx` en `~/Downloads/`.

⚠️ Esto sobrescribe cualquier dato ya cargado — hacer una copia de respaldo de la planilla con datos antes de regenerar.

## Salidas del dashboard

- KPIs: emisiones totales, Scope 1, Scope 2, capturas de sumideros, huella neta
- Distribución de emisiones por categoría (gráfico de torta)
- % de emisiones neutralizadas por sumideros propios (gauge)
- Detalle de Scope 1 por fuente
- Consumo eléctrico mensual y su equivalente en tCO2e
- Detalle de fermentación entérica por categoría animal y de suelos por cultivo
- Tabla resumen completa del inventario, exportable
