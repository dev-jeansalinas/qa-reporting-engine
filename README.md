# QA Evidence Generator

> Motor de generación de evidencias de QA que convierte datos en JSON en reportes automatizados de Excel y Word, optimizando el proceso de documentación de pruebas.

---

## Descripción

Este proyecto automatiza la creación de evidencias de pruebas (QA) a partir de archivos JSON, generando documentos estructurados en **Excel** y **Word** usando plantillas personalizadas.

Está diseñado para reducir el trabajo manual, mejorar la trazabilidad y estandarizar la documentación en procesos de testing.

---

## Problema que resuelve

En muchos equipos de QA:

- La documentación de evidencias es manual y repetitiva  
- Se pierde tiempo generando reportes  
- Hay inconsistencias en formatos  

Esta herramienta automatiza todo el proceso usando JSON como fuente única de datos.

---

## Funcionalidades

- ✅ Generación automática de Excel con formato profesional  
- ✅ Generación de evidencias en Word usando plantilla  
- ✅ Lectura de datos desde JSON  
- ✅ Configuración desacoplada (`config_proyecto.json`)  
- ✅ Inserción de logo en reportes  
- ✅ Formato condicional y validaciones en Excel  

---

## Tecnologías utilizadas

- Python  
- openpyxl (Excel)  
- python-docx (Word)  
- JSON  

---

## Estructura del proyecto

casos_de_pruebas/
│── assets/
│ ├── logo.png
│ └── plantilla_sinova.docx
│
│── data/
│ ├── casos.json
│ └── config_proyecto.json
│
│── output/
│
│── src/
│ ├── generator.py
│ └── utils.py
│── run.py
│── requirements.txt
│── README.md

## Instalación

```bash
git clone url_repositorio
cd qa-evidence-generator

python -m venv .venv
source .venv/bin/activate   # Linux / Mac
.venv\Scripts\activate      # Windows

pip install -r requirements.txt

## Ejecución

```bash
python run.py
```

## 📝 Estructura de Datos (JSON)

### config_proyecto.json

```json
{
    "modulo": "Nombre del Módulo",
    "cliente": "Nombre del Cliente",
    "tester": "Nombre del Tester",
    "descripcion": "Descripción del proyecto",
    "version": "Versión del módulo",
    "inicio_pruebas": "Fecha de inicio",
    "fin_pruebas": "Fecha de fin",
    "alojamiento": "Alojamiento del módulo"
}
```

### casos.json

```json
[
    {
        "nombre_caso": "Nombre del Caso de Prueba",
        "descripcion": "Descripción del caso de prueba",
        "pasos": "Pasos para ejecutar el caso de prueba",
        "datos": "Datos necesarios para ejecutar el caso de prueba",
        "resultado_esperado": "Resultado esperado del caso de prueba",
        "resultado_obtenido": "Resultado obtenido del caso de prueba"
    }
]
```

# Evidencias
<details>
    <summary>Ver Capturas de Reportes (Excel/Word)</summary>

## Excel
<img width="1825" height="439" alt="image" src="https://github.com/user-attachments/assets/10237d9c-9cc0-4993-8299-aa1d6757f6ec" />

## Word
<img width="815" height="1388" alt="image" src="https://github.com/user-attachments/assets/c7a90700-cc7a-4395-a6ef-075bf5fc9bf5" />
</details>

