# 💥 Análisis de Costos y Consumo de Explosivos por Voladura

Aplicación web desarrollada con **Streamlit** para el diseño y análisis técnico-económico de voladuras en minería a cielo abierto y subterránea.

**Método:** Langefors–Kihlström (1963) + Ash (1963)  
**Simulación:** Monte Carlo integrada  
**Autor:** Royer Casas

---

## 🚀 Instalación y Ejecución

### 1. Clonar el repositorio
```bash
git clone https://github.com/<tu-usuario>/<tu-repositorio>.git
cd <tu-repositorio>
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Ejecutar la aplicación
```bash
py -m streamlit run app.py
```
La app se abrirá automáticamente en tu navegador en `http://localhost:8501`

---

## 📋 Parámetros de Entrada

| Sección | Parámetros |
|---|---|
| Roca y Geomecánica | Densidad de la roca in situ (ρr) |
| Geometría del banco | Altura (H), ángulo (α), coeficientes Ks, Kj, Kt |
| Taladro y Explosivo | Diámetro (Dh), densidad explosivo (ρe), energía específica (e_expl) |
| Costos Unitarios | Ce, C_init, N_init, C_CD, Cp |
| Parámetros Operacionales | N° taladros (Nt), overhead (Fop) |
| Monte Carlo | N° de simulaciones (5,000 / 10,000 / 50,000) |

---

## 📊 Resultados y Visualizaciones

La aplicación genera automáticamente:

- **Resultados determinísticos:** Burden (B), espaciamiento (S), masa de explosivo (Qe), factor de carga (FC), RWS, RBS, costo total (CT) y costo unitario de voladura (CUV).
- **Histogramas y CDF** de las variables clave (CUV, FC, CT, B) con percentiles P5, P50, P95.
- **Diagrama de Tornado** con correlaciones de Pearson para identificar las variables más influyentes.
- **Scatter plots** de variables de entrada vs CUV.
- **Análisis de sensibilidad:** Burden vs diámetro, FC vs diámetro, mapa de contorno CUV vs Cp/Ce.
- **Comparación de explosivos comerciales:** ANFO, emulsión, mezclas.
- **Exportación a Excel** con hojas de datos de entrada, resultados, estadísticas MC y correlaciones.

---

## 📦 Dependencias

| Librería | Uso |
|---|---|
| `streamlit` | Framework de la aplicación web |
| `numpy` | Cálculos numéricos y simulación Monte Carlo |
| `scipy` | Distribuciones de probabilidad (truncnorm, pearsonr) |
| `matplotlib` | Generación de gráficos |
| `openpyxl` | Exportación a Excel (.xlsx) |
| `pandas` | Tablas de estadísticas en la interfaz |

---

## 🗂️ Estructura del Proyecto

```
├── app.py              # Aplicación Streamlit principal
├── requirements.txt    # Dependencias del proyecto
└── README.md           # Este archivo
```

---

## 📐 Fundamento Matemático

### Burden (Langefors–Kihlström)
```
B = 0.012 × √(2 × ρe / ρr) × Dh
```

### Masa de explosivo por taladro
```
Qe = π × r² × Lc × ρe × 1000   [kg]
```

### Factor de carga
```
FC = Qe / (B × S × H × ρr)   [kg/t]
```

### Costo unitario de voladura
```
CUV = CT / Ttotal   [USD/t]
CT  = Nt × (Ce_t + Cp_t) × (1 + Fop/100)
```

### RWS y RBS (energía relativa)
```
RWS = (e_expl / e_ANFO) × 100   [%]   (e_ANFO = 3.85 MJ/kg)
RBS = RWS × (ρe / ρ_ANFO)       [%]   (ρ_ANFO = 0.80 g/cm³)
```

---

## 📄 Licencia

Proyecto académico — Ingeniería de Minas.
