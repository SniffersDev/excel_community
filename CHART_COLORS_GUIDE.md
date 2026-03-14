# 🎨 Estrategia de Colores por Tipo de Gráfico

## Resumen Ejecutivo

Se implementó un sistema completo de colores para todos los tipos de gráficos de la librería `excel_community`, garantizando visualizaciones profesionales, distinguibles y estéticamente agradables.

---

## 📊 Gráficos de Series (Column, Bar, Line, Area, Scatter)

### Paleta de Colores (12 colores)
```
1.  #4472C4 - Blue (Azul)
2.  #ED7D31 - Orange (Naranja)
3.  #70AD47 - Green (Verde)
4.  #FFC000 - Gold (Dorado)
5.  #5B9BD5 - Light Blue (Azul Claro)
6.  #C5504B - Red (Rojo)
7.  #8064A2 - Purple (Púrpura)
8.  #4BACC6 - Cyan (Cian)
9.  #9BBB59 - Olive (Oliva)
10. #F79646 - Light Orange (Naranja Claro)
11. #17B897 - Teal (Verde Azulado)
12. #E83352 - Crimson (Carmesí)
```

### Rotación Automática
Si hay más de 12 series, los colores se reutilizan usando `i % 12`.

---

## 1️⃣ Column Chart (Gráfico de Columnas)

**Características:**
- ✅ Colores sólidos (100% opacidad)
- ✅ Bordes del mismo color que el relleno
- ✅ Grosor de borde: 9525 EMUs (fino)

**Uso:**
```dart
ColumnChart(
  title: "Sales by Quarter",
  series: series,  // Cada serie tendrá un color diferente
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Barras verticales con colores sólidos distinguibles
- Ideal para comparar categorías

---

## 2️⃣ Bar Chart (Gráfico de Barras Horizontales)

**Características:**
- ✅ Idéntico a Column Chart pero horizontal
- ✅ Colores sólidos (100% opacidad)
- ✅ Bordes del mismo color

**Uso:**
```dart
BarChart(
  title: "Revenue by Product",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Barras horizontales con colores sólidos
- Mejor para labels largos

---

## 3️⃣ Line Chart (Gráfico de Líneas)

**Características:**
- ✅ Líneas gruesas: 28575 EMUs
- ✅ Marcadores circulares (tamaño 5)
- ✅ Marcadores del mismo color que la línea
- ✅ Sin transparencia (100% opacidad)

**Uso:**
```dart
LineChart(
  title: "Trends Over Time",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Líneas coloridas con puntos circulares
- Ideal para tendencias temporales

---

## 4️⃣ Area Chart (Gráfico de Áreas)

**Características:**
- ✅ Líneas: 90% opacidad (alpha: 90000)
- ✅ Relleno: 50% opacidad (alpha: 50000)
- ✅ Grosor de línea: 28575 EMUs (grueso)

**Uso:**
```dart
AreaChart(
  title: "Market Share Evolution",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Áreas semitransparentes que permiten ver superposiciones
- Líneas más opacas para definir contornos
- Perfecto para mostrar acumulación o partes de un todo

---

## 5️⃣ Scatter Chart (Gráfico de Dispersión)

**Características:**
- ✅ Puntos circulares (tamaño 7)
- ✅ Relleno sólido del color de la serie
- ✅ Borde blanco (#FFFFFF) grosor 9525
- ✅ Sin transparencia

**Uso:**
```dart
ScatterChart(
  title: "Correlation Analysis",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Puntos coloridos con halo blanco
- El borde blanco ayuda a distinguir puntos superpuestos
- Ideal para correlaciones y distribuciones

---

## 🥧 Gráficos Circulares (Pie, Doughnut)

### Paleta de Colores (20 colores)
```
1.  #4472C4 - Blue          11. #255E91 - Navy Blue
2.  #ED7D31 - Orange        12. #43682B - Dark Green
3.  #A5A5A5 - Gray          13. #C5504B - Red
4.  #FFC000 - Gold          14. #8064A2 - Purple
5.  #5B9BD5 - Light Blue    15. #4BACC6 - Cyan
6.  #70AD47 - Green         16. #F79646 - Light Orange
7.  #264478 - Dark Blue     17. #9BBB59 - Olive
8.  #9E480E - Brown         18. #E83352 - Crimson
9.  #636363 - Dark Gray     19. #17B897 - Teal
10. #997300 - Dark Gold     20. #FF6F61 - Coral
```

### Algoritmo de Asignación
1. Shuffle (aleatorizar) la paleta
2. Tomar los primeros N colores (N = número de segmentos)
3. Asignar uno por uno a cada segmento

**Resultado:** Cada ejecución genera combinaciones aleatorias SIN repetición

---

## 6️⃣ Pie Chart (Gráfico Circular)

**Características:**
- ✅ Colores aleatorios sin repetición
- ✅ Colores sólidos (100% opacidad)
- ✅ 20 colores disponibles
- ✅ Shuffle antes de asignar

**Uso:**
```dart
PieChart(
  title: "Market Share",
  series: [series[0]], // Solo una serie
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Cada segmento tiene un color vibrante diferente
- Colores randomizados cada vez
- Máximo 20 segmentos con colores únicos

---

## 7️⃣ Doughnut Chart (Gráfico de Dona)

**Características:**
- ✅ Idéntico a Pie Chart en colores
- ✅ Hueco central (holeSize: 50%)
- ✅ Colores aleatorios sin repetición

**Uso:**
```dart
DoughnutChart(
  title: "Budget Distribution",
  series: [series[0]], // Solo una serie
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Como Pie Chart pero con hueco central
- Mismo sistema de colores aleatorios

---

## 🕸️ Radar Chart (Gráfico de Radar)

### Paleta de Colores (8 colores)
```
1. #4472C4 - Blue
2. #ED7D31 - Orange
3. #70AD47 - Green
4. #FFC000 - Gold
5. #5B9BD5 - Light Blue
6. #C5504B - Red
7. #8064A2 - Purple
8. #4BACC6 - Cyan
```

---

## 8️⃣ Radar Chart - Filled (Rellenado)

**Características:**
- ✅ Líneas: 85% opacidad (alpha: 85000)
- ✅ Relleno: 45% opacidad (alpha: 45000)
- ✅ Grosor de línea: 28575 EMUs (grueso)

**Uso:**
```dart
RadarChart(
  title: "Skills Assessment",
  series: series,
  anchor: anchor,
  showLegend: true,
  filled: true, // ← IMPORTANTE
)
```

**Visual:**
- Áreas muy transparentes (45%) permite ver superposiciones
- Líneas más visibles (85%)
- Perfecto para comparar múltiples perfiles

---

## 9️⃣ Radar Chart - Lines (Solo Líneas)

**Características:**
- ✅ Líneas: 85% opacidad (alpha: 85000)
- ✅ Sin relleno
- ✅ Grosor de línea: 28575 EMUs

**Uso:**
```dart
RadarChart(
  title: "Performance Metrics",
  series: series,
  anchor: anchor,
  showLegend: true,
  filled: false, // ← IMPORTANTE
)
```

**Visual:**
- Solo contornos sin relleno
- Más limpio cuando hay muchas series superpuestas

---

## 📊 Comparativa de Transparencias

| Tipo de Gráfico | Líneas | Relleno | Marcadores | Bordes |
|-----------------|--------|---------|------------|--------|
| Column/Bar | - | 100% | - | 100% |
| Line | 100% | - | 100% | - |
| Area | 90% | 50% | - | - |
| Scatter | - | 100% | 100% | Blanco |
| Pie/Doughnut | - | 100% | - | - |
| Radar (filled) | 85% | 45% | - | - |
| Radar (lines) | 85% | - | - | - |

---

## 🎯 Principios de Diseño Aplicados

### 1. **Distinguibilidad**
- Colores suficientemente diferentes entre sí
- Evita confusión visual

### 2. **Transparencia Estratégica**
- **Gráficos de área/radar:** Transparencia para ver superposiciones
- **Gráficos de barras/puntos:** Sólidos para máxima claridad

### 3. **Consistencia**
- Misma paleta base para todos los gráficos de series
- Paleta extendida para gráficos circulares (más segmentos)

### 4. **Profesionalismo**
- Colores basados en paletas de Office/Excel
- No demasiado brillantes ni apagados

### 5. **Accesibilidad**
- Incluye variaciones de tono (claro/oscuro)
- Buenos contrastes

---

## 📝 Notas de Implementación

### Unidades de Medida
- **EMUs (English Metric Units):** 914400 EMUs = 1 inch
- **Grosor de línea delgada:** 9525 EMUs ≈ 0.75 pt
- **Grosor de línea gruesa:** 28575 EMUs ≈ 2.25 pt

### Valores de Alpha (Transparencia)
- **100% opaco:** Sin elemento `<a:alpha>`
- **90% opaco:** `alpha="90000"`
- **85% opaco:** `alpha="85000"`
- **50% opaco:** `alpha="50000"`
- **45% opaco:** `alpha="45000"`

El valor de alpha va de 0 a 100000 (0% a 100%)

---

## 🧪 Archivos de Prueba Generados

Ejecuta `dart run test_all_colors.dart` para generar:

1. `COLOR_TEST_column_chart.xlsx`
2. `COLOR_TEST_bar_chart.xlsx`
3. `COLOR_TEST_line_chart.xlsx`
4. `COLOR_TEST_area_chart.xlsx`
5. `COLOR_TEST_scatter_chart.xlsx`
6. `COLOR_TEST_pie_chart.xlsx`
7. `COLOR_TEST_doughnut_chart.xlsx`
8. `COLOR_TEST_radar_filled_chart.xlsx`
9. `COLOR_TEST_radar_lines_chart.xlsx`

**Todos** los gráficos incluyen:
- 3 series de datos (excepto circulares que usan 1)
- 6 categorías
- Leyenda activada
- Títulos descriptivos

---

## ✅ Verificación de Calidad

Abre los archivos generados en Microsoft Excel y verifica:

- ✅ Todos los colores son distintos y visibles
- ✅ Las transparencias funcionan correctamente
- ✅ Los bordes y marcadores se ven bien
- ✅ Las leyendas muestran los colores correctos
- ✅ No hay colores repetidos en el mismo gráfico
- ✅ Los gráficos circulares tienen variedad aleatoria

---

## 🚀 Uso Recomendado

### Column/Bar Charts
- Comparaciones entre categorías
- Datos discretos
- Cada categoría claramente separada

### Line Charts
- Tendencias temporales
- Datos continuos
- Cambios a lo largo del tiempo

### Area Charts
- Acumulación de valores
- Contribución de partes al total
- Cuando la superposición es importante

### Scatter Charts
- Correlaciones
- Distribuciones
- Relaciones entre variables

### Pie/Doughnut Charts
- Partes de un todo (porcentajes)
- Máximo 6-8 segmentos para legibilidad
- Cuando el total suma 100%

### Radar Charts
- Comparar múltiples métricas
- Perfiles multidimensionales
- Evaluaciones de competencias

---

## 📦 Conclusión

Cada tipo de gráfico tiene una estrategia de color optimizada para su caso de uso específico:

- **Gráficos de barras/columnas:** Sólidos y distinguibles
- **Gráficos de líneas:** Marcadores visibles
- **Gráficos de área/radar:** Transparentes para ver superposiciones
- **Gráficos circulares:** Aleatorios para variedad
- **Gráficos de dispersión:** Bordes blancos para separación

Este sistema garantiza visualizaciones profesionales y fáciles de interpretar en todos los casos.

---

**Fecha de implementación:** 14 de marzo de 2026  
**Librería:** excel_community  
**Archivo de implementación:** `lib/src/utilities/chart_xml_writer.dart`
