# IDEP — Instrumento Diagnóstico de Ecosistemas Productivos
**Aplicación Python / Streamlit**  
Investigación Doctoral en Innovación y Productividad Regional · Modelo de la Quinta Hélice

---

## ▶️ Cómo ejecutar

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Ejecutar la aplicación
```bash
streamlit run idep_app.py
```

La aplicación se abrirá automáticamente en su navegador en `http://localhost:8501`

---

## 📦 Estructura de la herramienta

| Paso | Sección |
|------|---------|
| 1 | Identificación del actor (datos personales y organización) |
| 2 | Tipo de actor según Quinta Hélice |
| 3 | Tipología específica (7 opciones por hélice) |
| 4 | Ecosistemas productivos regionales (máx. 3 de 15) |
| 5 | Estado actual: madurez, competitividad, articulación (6 preguntas con escala + 200 palabras) |
| 6 | Mapeo de actores, cadena de valor y gobernanza (6 preguntas con escala + 200 palabras) |
| 7 | Diagnóstico de cadenas productivas y necesidades (6 preguntas abiertas + 200 palabras) |

## 📊 Exportación Excel
Al finalizar el formulario se genera un archivo `.xlsx` con:
- **Hoja 1:** Respuesta completa formateada y con estilos
- **Hoja 2:** Datos planos listos para análisis cuantitativo y cualitativo

---

## 📚 Referencias bibliográficas principales (Q1/Q2)
- Carayannis, E. G., & Campbell, D. F. J. (2010). *Int. J. Social Ecology and Sustainable Development, 1*(1), 41–69.
- Jacobides, M. G. et al. (2018). *Strategic Management Journal, 39*(8), 2255–2276.
- Granstrand, O., & Holgersson, M. (2020). *Technovation, 90–91*, 102098.
- Gereffi, G., & Fernandez-Stark, K. (2016). Global Value Chain Analysis. CGGC.
- Stam, E. (2015). *European Planning Studies, 23*(6), 1759–1762.
- Autio, E. et al. (2018). *Strategic Entrepreneurship Journal, 12*(1), 72–95.
