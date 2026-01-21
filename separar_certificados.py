"""
Sistema para Separar Certificados PDF
=====================================
Este script toma un PDF con múltiples certificados (uno por página)
y los separa en PDFs individuales, renombrándolos según el nombre
que aparece en cada certificado o según una lista proporcionada.

Estructura de carpetas:
    entrada/  -> Colocar aquí los PDFs a procesar
    salida/   -> Aquí se guardarán los certificados separados

Uso:
    1. Simple (procesa todos los PDFs en 'entrada/'):
       python separar_certificados.py
    
    2. Archivo específico:
       python separar_certificados.py --archivo entrada/certificados.pdf
    
    3. Con lista de nombres (Excel/CSV):
       python separar_certificados.py --lista nombres.xlsx
    
    4. Ver patrones de búsqueda disponibles:
       python separar_certificados.py --ver-patrones
"""

import fitz  # PyMuPDF
import os
import sys
import re
import argparse
from pathlib import Path


# Configuración de carpetas
CARPETA_ENTRADA = "entrada"
CARPETA_SALIDA = "salida"
ARCHIVO_PATRONES = "patrones.txt"


def limpiar_carpeta_salida():
    """
    Elimina todos los archivos PDF de la carpeta de salida antes de procesar.
    """
    carpeta_salida = Path(__file__).parent / CARPETA_SALIDA
    
    if not carpeta_salida.exists():
        return 0
    
    archivos_pdf = list(carpeta_salida.glob("*.pdf"))
    eliminados = 0
    
    for archivo in archivos_pdf:
        try:
            archivo.unlink()
            eliminados += 1
        except Exception as e:
            print(f"⚠️  No se pudo eliminar {archivo.name}: {e}")
    
    if eliminados > 0:
        print(f"🗑️  Se eliminaron {eliminados} archivo(s) de la carpeta 'salida/'")
    
    return eliminados


def eliminar_pdf_entrada(ruta_pdf: str):
    """
    Elimina un PDF de la carpeta de entrada después de procesarlo.
    """
    try:
        archivo = Path(ruta_pdf)
        if archivo.exists():
            archivo.unlink()
            print(f"🗑️  Eliminado de entrada: {archivo.name}")
            return True
    except Exception as e:
        print(f"⚠️  No se pudo eliminar {ruta_pdf}: {e}")
    return False


def limpiar_nombre_archivo(nombre: str) -> str:
    """
    Limpia el nombre para que sea válido como nombre de archivo.
    Elimina caracteres especiales y espacios extra.
    """
    # Eliminar caracteres no válidos para nombres de archivo
    nombre_limpio = re.sub(r'[<>:"/\\|?*]', '', nombre)
    # Reemplazar múltiples espacios por uno solo
    nombre_limpio = re.sub(r'\s+', ' ', nombre_limpio)
    # Eliminar espacios al inicio y final
    nombre_limpio = nombre_limpio.strip()
    # Limitar longitud del nombre
    if len(nombre_limpio) > 100:
        nombre_limpio = nombre_limpio[:100]
    # Si el nombre está vacío, usar un nombre por defecto
    if not nombre_limpio:
        nombre_limpio = "certificado_sin_nombre"
    return nombre_limpio


def cargar_patrones(ruta_archivo: str = None) -> list:
    """
    Carga los patrones de búsqueda desde el archivo patrones.txt
    
    Returns:
        Lista de patrones regex
    """
    if ruta_archivo is None:
        ruta_archivo = Path(__file__).parent / ARCHIVO_PATRONES
    
    patrones = []
    
    if not os.path.exists(ruta_archivo):
        print(f"⚠️  Archivo de patrones no encontrado: {ruta_archivo}")
        print("   Usando patrones por defecto...")
        return [
            r"Se otorga el presente reconocimiento a:\s*\n?\s*(.+?)(?:\n|Por su)",
            r"[Oo]torga(?:do)? a:\s*(.+?)(?:\n|$)",
            r"[Cc]ertifica(?:do)? a:\s*(.+?)(?:\n|$)",
        ]
    
    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as f:
            for linea in f:
                linea = linea.strip()
                # Ignorar líneas vacías y comentarios
                if linea and not linea.startswith('#'):
                    patrones.append(linea)
    except Exception as e:
        print(f"⚠️  Error al leer patrones: {e}")
    
    return patrones


def extraer_nombre_de_pagina(pagina, patrones: list = None) -> str:
    """
    Extrae el nombre del destinatario del certificado de una página.
    
    Args:
        pagina: Objeto de página de PyMuPDF
        patrones: Lista de patrones regex para buscar el nombre
    
    Returns:
        Nombre extraído o None si no se encuentra
    """
    texto = pagina.get_text()
    
    if patrones is None:
        patrones = cargar_patrones()
    
    for patron_regex in patrones:
        try:
            match = re.search(patron_regex, texto, re.IGNORECASE | re.DOTALL)
            if match:
                nombre = match.group(1).strip()
                # Limpiar el nombre de saltos de línea y espacios extra
                nombre = re.sub(r'\s+', ' ', nombre)
                # Eliminar texto innecesario que pueda haberse capturado
                nombre = nombre.split('\n')[0].strip()
                if nombre and len(nombre) > 2:
                    return nombre
        except re.error as e:
            print(f"⚠️  Patrón inválido ignorado: {patron_regex} ({e})")
    
    return None


def cargar_lista_nombres(ruta_archivo: str) -> list:
    """
    Carga una lista de nombres desde un archivo Excel o CSV.
    
    Args:
        ruta_archivo: Ruta al archivo Excel (.xlsx, .xls) o CSV
    
    Returns:
        Lista de nombres
    """
    import pandas as pd
    
    extension = Path(ruta_archivo).suffix.lower()
    
    try:
        if extension in ['.xlsx', '.xls']:
            df = pd.read_excel(ruta_archivo, header=None)
        elif extension == '.csv':
            df = pd.read_csv(ruta_archivo, header=None)
        else:
            raise ValueError(f"Formato de archivo no soportado: {extension}")
        
        # Tomar la primera columna como lista de nombres
        nombres = df.iloc[:, 0].astype(str).tolist()
        # Filtrar valores vacíos o NaN
        nombres = [n for n in nombres if n and n.lower() != 'nan' and n.strip()]
        
        return nombres
    
    except Exception as e:
        print(f"Error al cargar el archivo de lista: {e}")
        return []


def separar_certificados(
    ruta_pdf: str,
    carpeta_salida: str = None,
    lista_nombres: list = None,
    patrones: list = None,
    prefijo: str = "",
    sufijo: str = ""
) -> dict:
    """
    Separa un PDF de múltiples certificados en archivos individuales.
    """
    # Validar que el archivo existe
    if not os.path.exists(ruta_pdf):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_pdf}")
    
    # Usar carpeta de salida por defecto
    if carpeta_salida is None:
        carpeta_salida = Path(__file__).parent / CARPETA_SALIDA
    
    carpeta_salida = Path(carpeta_salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    
    # Cargar patrones si no se proporcionan
    if patrones is None:
        patrones = cargar_patrones()
    
    # Abrir el PDF
    documento = fitz.open(ruta_pdf)
    total_paginas = len(documento)
    
    nombre_pdf = Path(ruta_pdf).name
    
    print(f"\n{'='*60}")
    print(f"📄 Archivo: {nombre_pdf}")
    print(f"📑 Total de páginas/certificados: {total_paginas}")
    print(f"📁 Carpeta de salida: {carpeta_salida}")
    print(f"🔍 Patrones cargados: {len(patrones)}")
    print(f"{'='*60}\n")
    
    # Validar lista de nombres si se proporciona
    if lista_nombres:
        if len(lista_nombres) < total_paginas:
            print(f"⚠️  Advertencia: La lista tiene {len(lista_nombres)} nombres pero hay {total_paginas} páginas.")
            print("    Se usará extracción automática para las páginas restantes.\n")
    
    resultados = {
        "exitosos": [],
        "errores": [],
        "total": total_paginas,
        "pdf_origen": nombre_pdf
    }
    
    # Procesar cada página
    for i in range(total_paginas):
        pagina = documento[i]
        numero = i + 1
        
        # Determinar el nombre del archivo
        if lista_nombres and i < len(lista_nombres):
            nombre = lista_nombres[i]
            origen = "lista"
        else:
            nombre = extraer_nombre_de_pagina(pagina, patrones)
            origen = "extraído"
        
        # Si no se pudo obtener un nombre, usar número de página
        if not nombre:
            nombre = f"certificado_{numero:03d}"
            origen = "generado"
        
        # Limpiar nombre para uso como archivo
        nombre_limpio = limpiar_nombre_archivo(nombre)
        nombre_archivo = f"{prefijo}{nombre_limpio}{sufijo}.pdf"
        ruta_salida = carpeta_salida / nombre_archivo
        
        # Manejar nombres duplicados
        contador = 1
        while ruta_salida.exists():
            nombre_archivo = f"{prefijo}{nombre_limpio}_{contador}{sufijo}.pdf"
            ruta_salida = carpeta_salida / nombre_archivo
            contador += 1
        
        try:
            # Crear nuevo PDF con esta página
            nuevo_pdf = fitz.open()
            nuevo_pdf.insert_pdf(documento, from_page=i, to_page=i)
            nuevo_pdf.save(str(ruta_salida))
            nuevo_pdf.close()
            
            estado = "✅" if origen != "generado" else "⚠️"
            print(f"{estado} [{numero}/{total_paginas}] {nombre_archivo}")
            print(f"   └─ Nombre: {nombre} ({origen})")
            
            resultados["exitosos"].append({
                "pagina": numero,
                "nombre": nombre,
                "archivo": str(ruta_salida),
                "origen": origen
            })
            
        except Exception as e:
            print(f"❌ [{numero}/{total_paginas}] Error al procesar página: {e}")
            resultados["errores"].append({
                "pagina": numero,
                "error": str(e)
            })
    
    documento.close()
    
    # Contar por origen
    extraidos = len([r for r in resultados["exitosos"] if r["origen"] == "extraído"])
    de_lista = len([r for r in resultados["exitosos"] if r["origen"] == "lista"])
    generados = len([r for r in resultados["exitosos"] if r["origen"] == "generado"])
    
    # Resumen final
    print(f"\n{'='*60}")
    print(f"📊 RESUMEN - {nombre_pdf}")
    print(f"{'='*60}")
    print(f"✅ Exitosos: {len(resultados['exitosos'])}/{total_paginas}")
    if extraidos > 0:
        print(f"   └─ Nombres extraídos del PDF: {extraidos}")
    if de_lista > 0:
        print(f"   └─ Nombres de lista: {de_lista}")
    if generados > 0:
        print(f"   └─ Nombres generados (no encontrados): {generados}")
    print(f"❌ Errores: {len(resultados['errores'])}/{total_paginas}")
    print(f"📁 Archivos guardados en: {carpeta_salida}")
    print(f"{'='*60}\n")
    
    return resultados


def procesar_carpeta_entrada(
    lista_nombres: list = None,
    prefijo: str = "",
    sufijo: str = "",
    limpiar_salida: bool = True,
    eliminar_entrada: bool = True
) -> list:
    """
    Procesa todos los PDFs en la carpeta de entrada.
    
    Args:
        lista_nombres: Lista opcional de nombres para renombrar
        prefijo: Prefijo para agregar al nombre de cada archivo
        sufijo: Sufijo para agregar al nombre de cada archivo
        limpiar_salida: Si True, elimina los PDFs existentes en salida antes de procesar
        eliminar_entrada: Si True, elimina los PDFs de entrada después de procesarlos
    """
    carpeta_entrada = Path(__file__).parent / CARPETA_ENTRADA
    carpeta_salida = Path(__file__).parent / CARPETA_SALIDA
    
    if not carpeta_entrada.exists():
        carpeta_entrada.mkdir(parents=True, exist_ok=True)
        print(f"📁 Se creó la carpeta 'entrada/'. Coloca tus PDFs ahí y ejecuta de nuevo.")
        return []
    
    archivos_pdf = list(carpeta_entrada.glob("*.pdf"))
    
    if not archivos_pdf:
        print(f"⚠️  No se encontraron archivos PDF en: {carpeta_entrada}")
        print("   Coloca tus PDFs en la carpeta 'entrada/' y ejecuta de nuevo.")
        return []
    
    print(f"\n🔍 Encontrados {len(archivos_pdf)} archivo(s) PDF en 'entrada/'")
    
    # Limpiar carpeta de salida antes de procesar
    if limpiar_salida:
        limpiar_carpeta_salida()
    
    todos_resultados = []
    patrones = cargar_patrones()
    pdfs_procesados = []  # Lista de PDFs procesados exitosamente
    
    for pdf in archivos_pdf:
        try:
            resultado = separar_certificados(
                ruta_pdf=str(pdf),
                carpeta_salida=str(carpeta_salida),
                lista_nombres=lista_nombres,
                patrones=patrones,
                prefijo=prefijo,
                sufijo=sufijo
            )
            todos_resultados.append(resultado)
            
            # Si no hubo errores, marcar para eliminar
            if not resultado["errores"]:
                pdfs_procesados.append(str(pdf))
                
        except Exception as e:
            print(f"❌ Error procesando {pdf.name}: {e}")
    
    # Eliminar PDFs de entrada después de procesar exitosamente
    if eliminar_entrada and pdfs_procesados:
        print(f"\n🗑️  Limpiando carpeta de entrada...")
        for pdf_path in pdfs_procesados:
            eliminar_pdf_entrada(pdf_path)
    
    return todos_resultados


def mostrar_patrones():
    """Muestra los patrones de búsqueda actuales."""
    patrones = cargar_patrones()
    
    print(f"\n{'='*60}")
    print("🔍 PATRONES DE BÚSQUEDA CONFIGURADOS")
    print(f"{'='*60}")
    print(f"📄 Archivo: {ARCHIVO_PATRONES}\n")
    
    for i, patron in enumerate(patrones, 1):
        print(f"  {i}. {patron}")
    
    print(f"\n{'='*60}")
    print("💡 Para agregar nuevos patrones, edita el archivo 'patrones.txt'")
    print(f"{'='*60}\n")


def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(
        description="Separar PDFs de certificados en archivos individuales",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python separar_certificados.py                           # Procesa todos los PDFs en 'entrada/'
  python separar_certificados.py --archivo mi_archivo.pdf  # Procesa un archivo específico
  python separar_certificados.py --lista nombres.xlsx      # Usa nombres de un Excel
  python separar_certificados.py --ver-patrones            # Muestra los patrones configurados
        """
    )
    
    parser.add_argument(
        "--archivo", "-a",
        help="Ruta a un archivo PDF específico (si no se especifica, procesa la carpeta 'entrada/')",
        default=None
    )
    
    parser.add_argument(
        "--lista", "-l",
        help="Archivo Excel o CSV con lista de nombres (en orden)",
        default=None
    )
    
    parser.add_argument(
        "--salida", "-o",
        help="Carpeta donde guardar los certificados separados (por defecto: 'salida/')",
        default=None
    )
    
    parser.add_argument(
        "--prefijo",
        help="Prefijo para agregar al nombre de cada archivo",
        default=""
    )
    
    parser.add_argument(
        "--sufijo",
        help="Sufijo para agregar al nombre de cada archivo (antes de .pdf)",
        default=""
    )
    
    parser.add_argument(
        "--ver-patrones",
        action="store_true",
        help="Muestra los patrones de búsqueda configurados"
    )
    
    parser.add_argument(
        "--no-limpiar",
        action="store_true",
        help="No eliminar archivos de salida antes de procesar"
    )
    
    parser.add_argument(
        "--no-borrar-entrada",
        action="store_true",
        help="No eliminar PDFs de entrada después de procesarlos"
    )
    
    args = parser.parse_args()
    
    # Mostrar patrones si se solicita
    if args.ver_patrones:
        mostrar_patrones()
        return
    
    # Cargar lista de nombres si se proporciona
    lista_nombres = None
    if args.lista:
        if not os.path.exists(args.lista):
            print(f"❌ Error: No se encontró el archivo de lista: {args.lista}")
            sys.exit(1)
        lista_nombres = cargar_lista_nombres(args.lista)
        if lista_nombres:
            print(f"📋 Lista cargada con {len(lista_nombres)} nombres")
        else:
            print("⚠️  No se pudieron cargar nombres de la lista. Se usará extracción automática.")
    
    # Procesar archivo específico o carpeta
    try:
        if args.archivo:
            resultado = separar_certificados(
                ruta_pdf=args.archivo,
                carpeta_salida=args.salida,
                lista_nombres=lista_nombres,
                prefijo=args.prefijo,
                sufijo=args.sufijo
            )
            if resultado["errores"]:
                sys.exit(1)
        else:
            resultados = procesar_carpeta_entrada(
                lista_nombres=lista_nombres,
                prefijo=args.prefijo,
                sufijo=args.sufijo,
                limpiar_salida=not args.no_limpiar,
                eliminar_entrada=not args.no_borrar_entrada
            )
            if not resultados:
                sys.exit(1)
        
        sys.exit(0)
        
    except Exception as e:
        print(f"❌ Error fatal: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
