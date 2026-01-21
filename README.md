# 📄 Sistema Separador de Certificados PDF

Sistema en Python para separar un PDF con múltiples certificados en archivos individuales, renombrándolos automáticamente según el nombre del destinatario.

## ✨ Características

- 🔄 **Separación automática**: Divide un PDF con múltiples páginas en archivos individuales
- 🔍 **Extracción inteligente de nombres**: Detecta automáticamente el nombre del destinatario en cada certificado
- 📝 **Patrones personalizables**: Archivo `patrones.txt` editable para adaptarse a diferentes formatos de certificados
- 📋 **Soporte para listas**: Opcionalmente usa un archivo Excel/CSV con los nombres en orden
- 📁 **Organización simple**: Carpetas `entrada/` y `salida/` para facilitar el proceso

## 📁 Estructura del Proyecto

```
SISTEMA_CERTIFICADOS/
├── entrada/              # Coloca aquí los PDFs a procesar
├── salida/               # Aquí se guardarán los certificados separados
├── patrones.txt          # Patrones de búsqueda configurables
├── separar_certificados.py
├── requirements.txt
└── README.md
```

## 🚀 Instalación

### 1. Crear entorno virtual (recomendado)

```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno virtual (Windows)
.\venv\Scripts\activate

# Activar entorno virtual (Linux/Mac)
source venv/bin/activate
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

## 📖 Uso

### Modo básico (procesar todos los PDFs en `entrada/`)

1. Coloca tu(s) PDF(s) con certificados en la carpeta `entrada/`
2. Ejecuta:

```bash
python separar_certificados.py
```

3. Los certificados separados estarán en la carpeta `salida/`

### Procesar un archivo específico

```bash
python separar_certificados.py --archivo ruta/al/archivo.pdf
```

### Usar lista de nombres (Excel/CSV)

Si prefieres proporcionar los nombres en orden en lugar de extraerlos automáticamente:

```bash
python separar_certificados.py --lista nombres.xlsx
```

El archivo Excel/CSV debe tener los nombres en la primera columna, uno por fila.

### Ver patrones configurados

```bash
python separar_certificados.py --ver-patrones
```

### Opciones adicionales

| Opción | Descripción |
|--------|-------------|
| `--archivo`, `-a` | Procesar un archivo PDF específico |
| `--lista`, `-l` | Archivo Excel/CSV con lista de nombres |
| `--salida`, `-o` | Carpeta de salida personalizada |
| `--prefijo` | Agregar prefijo al nombre de cada archivo |
| `--sufijo` | Agregar sufijo al nombre de cada archivo |
| `--ver-patrones` | Mostrar patrones de búsqueda configurados |

### Ejemplos

```bash
# Procesar todos los PDFs en entrada/
python separar_certificados.py

# Procesar archivo específico
python separar_certificados.py --archivo entrada/certificados.pdf

# Con nombres desde Excel
python separar_certificados.py --lista participantes.xlsx

# Con prefijo en los archivos
python separar_certificados.py --prefijo "CERT_2025_"

# Combinación de opciones
python separar_certificados.py --archivo evento.pdf --prefijo "PHYLIVE_" --sufijo "_2025"
```

## ⚙️ Configurar Patrones de Búsqueda

El archivo `patrones.txt` contiene los patrones regex que el sistema usa para encontrar el nombre del destinatario en cada certificado.

### Formato del archivo

```txt
# Esto es un comentario (las líneas con # se ignoran)
Se otorga el presente reconocimiento a:\s*(.+?)(?:\n|Por su)
[Oo]torga(?:do)? a:\s*(.+?)(?:\n|$)
```

### Cómo agregar un nuevo patrón

1. Abre el PDF y observa qué texto aparece justo antes del nombre
2. Crea un patrón regex donde `(.+?)` captura el nombre
3. Agrega el patrón a `patrones.txt`

### Ejemplos de patrones comunes

| Texto en el certificado | Patrón a agregar |
|------------------------|------------------|
| "Se otorga a: **Juan Pérez**" | `Se otorga a:\s*(.+?)(?:\n\|$)` |
| "Certificamos que **María García**" | `Certificamos que\s+(.+?)(?:\n\|$)` |
| "A nombre de: **Carlos López**" | `A nombre de:\s*(.+?)(?:\n\|$)` |

## 🔧 Solución de Problemas

### Los nombres no se extraen correctamente

1. Ejecuta `python separar_certificados.py --ver-patrones` para ver los patrones actuales
2. Abre el PDF y observa el texto exacto que precede al nombre
3. Agrega un nuevo patrón a `patrones.txt`

### El script muestra "certificado_001.pdf" en lugar del nombre

Esto significa que ningún patrón coincidió. Revisa:
- El texto exacto del certificado (puede haber caracteres especiales)
- Agrega un patrón personalizado en `patrones.txt`

### Error de codificación con caracteres especiales

Asegúrate de que `patrones.txt` esté guardado con codificación UTF-8.

## 📋 Requisitos

- Python 3.8 o superior
- PyMuPDF (fitz)
- pandas (para listas Excel/CSV)
- openpyxl (para archivos .xlsx)

## 📄 Licencia

Este proyecto es de uso libre. Puedes modificarlo y distribuirlo como desees.

---

Desarrollado con ❤️ para facilitar la gestión de certificados
