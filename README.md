# Organizador EPLA - Gestión de Archivos de Secretaría

<div align="center">
  <img src="logo.png" alt="Logo Escuelas Profesionales Luis Amigó" width="150" height="150">
</div>

Una aplicación de escritorio desarrollada en Python con interfaz gráfica para automatizar la organización de archivos de estudiantes basándose en datos de un archivo Excel.

## 📋 Descripción

**Organizador EPLA** es una herramienta diseñada para las Escuelas Profesionales Luis Amigó que permite organizar automáticamente archivos de estudiantes descargándolos desde URLs (incluyendo Google Drive) y organizándolos en carpetas estructuradas según los datos del archivo Excel proporcionado.

## ✨ Características Principales

### Funcionalidades Técnicas

- **Interfaz Gráfica Intuitiva**: Desarrollada con tkinter para una experiencia de usuario amigable
- **Procesamiento de Excel**: Lee archivos Excel (.xlsx, .xls) utilizando pandas para extraer información de estudiantes
- **Descarga Automática**: Descarga archivos desde URLs, con soporte especial para enlaces de Google Drive
- **Organización Automática**: Crea estructura de carpetas basada en nombre y apellidos de estudiantes
- **Procesamiento Asíncrono**: Utiliza threading para evitar bloqueos en la interfaz durante el procesamiento
- **Manejo de Errores**: Sistema robusto de manejo de errores con retroalimentación al usuario
- **Detección Inteligente de PDFs**: Verifica tipos de contenido y garantiza la descarga correcta de archivos PDF
- **Prevención de Duplicados**: Evita descargar archivos que ya existen en el destino

### Características Estéticas

- **Logo Institucional**: Incluye el logo de EPLA (150x150 píxeles) en la interfaz principal
- **Diseño Profesional**: Interfaz limpia y profesional acorde a una institución educativa
- **Ventana Fija**: Tamaño de ventana optimizado (600x650 píxeles) no redimensionable para consistencia visual
- **Barra de Progreso**: Indicador visual del progreso del procesamiento con porcentajes
- **Mensajes Informativos**: Retroalimentación clara sobre el estado de las operaciones
- **Botones Estilizados**: Botón principal con estilo "Accent" y fuente en negrita
- **Espaciado Consistente**: Padding y márgenes optimizados para una experiencia visual agradable

## 🛠️ Tecnologías Utilizadas

### Dependencias Principales

- **Python 3.x**: Lenguaje de programación principal
- **tkinter**: Framework nativo para la interfaz gráfica
- **PIL (Pillow)**: Procesamiento de imágenes para el logo (redimensionamiento con LANCZOS)
- **pandas**: Manipulación y análisis de datos Excel
- **requests**: Realización de peticiones HTTP para descargas con streaming
- **pathlib**: Manejo moderno de rutas de archivos
- **threading**: Procesamiento asíncrono para mantener la GUI responsiva
- **re**: Expresiones regulares para procesamiento de URLs de Google Drive
- **urllib.parse**: Análisis y manipulación de URLs

### Empaquetado

- **PyInstaller**: Utilizado para crear el ejecutable independiente
- **UPX**: Compresión del ejecutable para reducir tamaño
- Incluye todas las dependencias necesarias y el logo en el bundle final
- Configurado para ejecutarse sin consola (modo windowed)

## 📁 Estructura del Proyecto

```
Organizador secretaría/
├── organizador_epla.py          # Código principal de la aplicación
├── logo.png                     # Logo institucional (150x150px)
├── organizador_epla.spec        # Configuración de PyInstaller
├── README.md                    # Documentación del proyecto
├── Organizador secretaría.code-workspace  # Configuración de VS Code
└── build/                       # Archivos de construcción
    └── organizador_epla/        # Ejecutable y dependencias
        ├── organizador_epla.pkg # Ejecutable principal
        ├── base_library.zip     # Bibliotecas base de Python
        ├── PYZ-00.pyz          # Código Python comprimido
        └── localpycs/          # Módulos Python compilados
```

## 🚀 Instalación y Uso

### Requisitos del Sistema

- **Sistema Operativo**: Windows, macOS, Linux
- **RAM**: 4GB mínimo recomendado
- **Espacio en disco**: 500MB para la aplicación y archivos temporales
- **Conexión a internet**: Necesaria para descargar archivos desde URLs

### Ejecución

1. **Ejecutable**: Utilizar el archivo ejecutable generado en `build/organizador_epla/organizador_epla.pkg`
2. **Desde código fuente**: 
   ```bash
   python organizador_epla.py
   ```

### Flujo de Trabajo

1. **Seleccionar Carpeta de Destino**: Elegir dónde se organizarán los archivos
2. **Seleccionar Archivo Excel**: Cargar el archivo con los datos de estudiantes
3. **Iniciar Procesamiento**: La aplicación automáticamente:
   - Lee los datos del Excel
   - Crea carpetas por estudiante (formato: "Apellido1 Apellido2 Nombre")
   - Descarga archivos desde las URLs de las columnas J-Q
   - Organiza todo según la estructura definida
   - Muestra progreso en tiempo real

## 📊 Configuración del Excel

### Estructura Requerida

El archivo Excel debe contener las siguientes columnas **exactas**:

| Columna | Nombre de Columna | Descripción |
|---------|-------------------|-------------|
| C | "Nombre del alumno/a" | Nombre del estudiante |
| D | "Primer apellido del alumno/a" | Primer apellido |
| E | "Segundo apellido del alumno/a" | Segundo apellido |
| J-Q | URLs de archivos | Enlaces a documentos PDF (columnas 9-16) |

### Formato de URLs Soportadas

- **URLs directas**: Enlaces directos a archivos PDF
- **Google Drive**: 
  - `https://drive.google.com/file/d/[ID]/view`
  - `https://drive.google.com/open?id=[ID]`
  - `https://drive.google.com/uc?id=[ID]`
- **Cualquier URL pública**: Accesible sin autenticación

### Nomenclatura de Archivos

Los archivos descargados se nombran según la columna de origen:
- Ejemplo: Si está en la columna "Documento de identidad", el archivo se guardará como `Documento de identidad.pdf`

## 🔧 Características Técnicas Avanzadas

### Manejo Inteligente de Google Drive

```python
def get_google_drive_direct_download_url(self, gd_url):
    """Convierte URLs de visualización de Google Drive en enlaces de descarga directa"""
    match = re.search(r'drive\.google\.com/(?:file/d/|open\?id=|uc\?id=)([a-zA-Z0-9_-]+)', gd_url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return None
```

### Procesamiento Asíncrono

- **Threading**: Operaciones de descarga en hilo separado para mantener GUI responsiva
- **Callbacks**: Actualización de progreso mediante `self.after()` para thread-safety
- **Timeout**: 15 segundos de timeout para descargas HTTP

### Gestión de Archivos

- **Sanitización**: Limpieza de nombres de carpetas eliminando caracteres especiales
- **Detección de tipos**: Verificación de Content-Type para asegurar archivos PDF
- **Streaming**: Descarga por chunks (8192 bytes) para optimizar memoria
- **Verificación de existencia**: Evita descargas duplicadas

### Validación y Manejo de Errores

- **Validación de columnas**: Verifica existencia de columnas requeridas
- **Datos incompletos**: Omite filas con información faltante
- **Errores de red**: Manejo robusto de timeouts y errores HTTP
- **Feedback visual**: Mensajes de error específicos para cada tipo de problema

## 📈 Rendimiento

- **Concurrencia**: Descarga un archivo por vez para evitar saturación
- **Memoria**: Streaming de archivos para manejar documentos grandes sin consumir RAM excesiva
- **Escalabilidad**: Capaz de procesar cientos de estudiantes y archivos
- **Optimización**: UPX compression reduce el tamaño del ejecutable

## 🛡️ Seguridad y Confiabilidad

- **Validación de entrada**: Verificación de integridad de archivos Excel
- **Sanitización de nombres**: Prevención de inyección de path en nombres de archivo
- **Timeout de red**: Prevención de colgado en descargas lentas
- **Backup automático**: Preserva archivos existentes evitando sobrescrituras
- **Logging**: Registro detallado de operaciones para debugging

## 🎨 Interfaz de Usuario

### Componentes Principales

1. **Header**: Logo institucional redimensionado con manejo de errores
2. **Paso 1**: Selector de carpeta de destino con campo de solo lectura
3. **Paso 2**: Selector de archivo Excel con filtros de tipo
4. **Botón principal**: Estilo destacado para iniciar procesamiento
5. **Barra de progreso**: Indicador determinista con porcentajes
6. **Estado**: Label informativo del progreso actual

### Responsividad

- **Estados de botones**: Deshabilitación durante procesamiento
- **Feedback inmediato**: Actualización en tiempo real del progreso
- **Prevención de ejecución múltiple**: Control de estado para evitar procesos duplicados

## 💡 Casos de Uso

### Escenario Principal
- **Secretaría académica** necesita organizar documentos de matriculación
- **Entrada**: Excel con datos de estudiantes y URLs de documentos
- **Salida**: Estructura de carpetas organizada con todos los documentos descargados

### Beneficios
- **Automatización**: Elimina trabajo manual repetitivo
- **Consistencia**: Nomenclatura uniforme de carpetas y archivos
- **Eficiencia**: Procesamiento masivo con mínima intervención
- **Trazabilidad**: Logs detallados de todo el proceso

## 📞 Soporte Técnico

### Requisitos de Datos
- Excel debe tener las columnas exactas especificadas
- URLs deben ser públicamente accesibles
- Archivos deben ser PDFs válidos

### Solución de Problemas
- **Logo no encontrado**: Verificar que `logo.png` esté en la carpeta del ejecutable
- **Error de columnas**: Verificar nombres exactos en el Excel
- **Fallos de descarga**: Verificar conectividad y validez de URLs
- **Memoria insuficiente**: Procesar archivos en lotes más pequeños

---

*Desarrollado específicamente para las **Escuelas Profesionales Luis Amigó** para optimizar los procesos administrativos de secretaría académica.*

## 🏷️ Versión

**Versión actual**: 1.0  
**Fecha de última actualización**: Mayo 2025  
**Compatibilidad**: Python 3.x, Windows/macOS/Linux
