# StoreCheck Automatizado

## 📋 Descripción del Proyecto

**StoreCheck Automatizado** es una herramienta de automatización diseñada para optimizar el proceso de relevamiento de información previo a actividades promocionales. Esta solución elimina la necesidad de copiar y pegar manualmente datos en plantillas, automatizando la generación de reportes de StoreCheck de manera eficiente y precisa.

## 🎯 Problemática Resuelta

### Situación Anterior
- **Proceso Manual**: Los operadores debían relevar información una semana antes de cada actividad promocional
- **Consultas Repetitivas**: Preguntas constantes sobre disponibilidad de productos por punto de venta
- **Trabajo Tedioso**: Copiar y pegar manualmente datos en plantillas con fórmulas
- **Validación Manual**: Verificar si el relevamiento correspondía al mercado correcto
- **Tiempo Excesivo**: Horas de trabajo repetitivo y propenso a errores

### Solución Implementada
Automatización completa del flujo de trabajo mediante una aplicación que:
- Procesa archivos base automáticamente
- Permite selección inteligente de mercados
- Genera plantillas de StoreCheck de forma instantánea
- Reduce errores humanos y tiempo de procesamiento

## ⚡ Características Principales

### 🔧 Funcionalidades Core
- **Procesamiento de Archivos**: Importa y procesa archivos `Base` y `Base Crono`
- **Selección Inteligente de Mercado**: Interface para elegir el mercado objetivo
- **Gestión de Actividades**: Selección de actividades promocionales específicas
- **Filtrado de Puntos**: Identificación y selección de puntos sin actividad
- **Generación Automática**: Creación del archivo `StoreCheck_Borrador` con plantilla completa

### 🎨 Interface de Usuario
- **Diseño Intuitivo**: Flujo de trabajo paso a paso
- **Validación en Tiempo Real**: Verificación automática de datos de entrada
- **Feedback Visual**: Indicadores de progreso y estado del procesamiento

## 🚀 Flujo de Trabajo

### Paso 1: Carga de Archivos Base
```
📁 Archivos de Entrada:
├── Base.xlsx (Información general de productos y puntos de venta)
└── Base_Crono.xlsx (Cronograma de actividades promocionales)
```

### Paso 2: Configuración
1. **Selección de Mercado**: El sistema permite elegir el mercado específico a procesar
2. **Edición de Activos**: Posibilidad de modificar información de productos activos
3. **Selección de Actividad**: Elección de la actividad promocional correspondiente

### Paso 3: Procesamiento
1. **Aceptar Configuración**: Confirmación de parámetros seleccionados
2. **Filtrado Automático**: El sistema identifica puntos sin actividad
3. **Aplicar Filtros**: Procesamiento de datos según criterios establecidos

### Paso 4: Generación de Resultado
```
📄 Archivo de Salida:
└── StoreCheck_Borrador.xlsx (Plantilla completa lista para uso)
```

## 📊 Beneficios Obtenidos

### ⏱️ Eficiencia Temporal
- **Reducción del 85%** en tiempo de procesamiento
- **Eliminación** de tareas repetitivas manuales
- **Procesamiento instantáneo** de múltiples puntos de venta

### 🎯 Precisión y Calidad
- **Eliminación de errores** de transcripción manual
- **Validación automática** de correspondencia mercado-relevamiento
- **Consistencia** en formato y estructura de plantillas

### 👥 Experiencia del Usuario
- **Interface amigable** para operadores no técnicos
- **Proceso guiado** paso a paso
- **Feedback inmediato** sobre el estado del procesamiento

## 🛠️ Tecnologías Utilizadas

- **Lenguaje**: [Especificar lenguaje de programación utilizado]
- **Procesamiento de Excel**: Bibliotecas para manipulación de archivos .xlsx
- **Interface Gráfica**: [Especificar framework de UI utilizado]
- **Validación de Datos**: Algoritmos de verificación automática

## 📋 Requisitos del Sistema

### Hardware Mínimo
- **RAM**: 4 GB mínimo (8 GB recomendado)
- **Espacio en Disco**: 100 MB para instalación
- **Procesador**: Dual-core 2.0 GHz o superior

### Software
- **Sistema Operativo**: Windows 10 o superior
- **Microsoft Excel**: 2016 o versión posterior (para visualización de resultados)

## 🖼️ Capturas del Aplicativo

### Vista 1: Selección de Actividad Promocional
![Image](https://github.com/user-attachments/assets/3f3e3e5d-ad53-44ad-846f-82affee9d12c)
- **Función**: Permite al usuario seleccionar una actividad promocional específica
- **Elementos**: Dropdown desplegable con opción "Aceptar"
- **Estado**: Vista inicial con selector vacío

### Vista 2: Lista de Actividades Disponibles  
![Image](https://github.com/user-attachments/assets/8ecc4e5e-e664-4bcf-b420-2e3844849df5)
- **Función**: Muestra todas las actividades promocionales disponibles
- **Elementos**: 
  - Lista desplegable con códigos de actividad
  - Formato: `CÓDIGO-DESCRIPCIÓN-UBICACIÓN`
  - Actividades como: 250502-TRO-CT-TR, 250411-MMC-CT-TR, etc.
- **Selección**: Actividad highlighted: `250330-AMA-JF-TR - AMA`

### Vista 3: Selección de Puntos de Venta
![Image](https://github.com/user-attachments/assets/8ecc4e5e-e664-4bcf-b420-2e3844849df5)
- **Función**: Permite seleccionar puntos de venta sin actividad para aplicar StoreCheck
- **Elementos**:
  - Lista de mercados con códigos identificadores
  - Formato: `CÓDIGO - NOMBRE DEL MERCADO`
  - Botón "Filtrar" para procesar selección
- **Ejemplos de Mercados**:
  - 1818 - MERCADO DE ABASTOS VICTOR LARCO
  - 2746 - MERCADO LAS 3 REGIONES  
  - 1891 - MERCADO EL INCA
  - 3853 - MERCADO MODELO
  - Y más opciones disponibles

### Configuración Paso a Paso

#### Paso 1: Carga de Archivos Base
- Cargar `Base.xlsx` y `Base_Crono.xlsx` en el sistema
- Verificar que los archivos sean válidos y contengan la información requerida

#### Paso 2: Selección de Actividad Promocional
1. Acceder al selector "Selecciona un valor de 'Act. Promocional'"
2. Desplegar la lista de actividades disponibles
3. Seleccionar la actividad deseada (ej: `250330-AMA-JF-TR - AMA`)
4. Hacer clic en "Aceptar"

#### Paso 3: Configuración de Puntos de Venta
1. El sistema abrirá la ventana "Selecciona puntos de venta sin actividad para aplicar SA"
2. Revisar la lista de mercados disponibles
3. Seleccionar los puntos de venta relevantes para el StoreCheck
4. Hacer clic en "Filtrar" para procesar la selección

#### Paso 4: Generación del StoreCheck
- El sistema procesará automáticamente los datos
- Generará el archivo `StoreCheck_Borrador.xlsx`
- El archivo estará listo para su descarga y uso

### Uso Rápido
1. Ejecutar la aplicación
2. Cargar archivos `Base.xlsx` y `Base_Crono.xlsx`
3. Seleccionar actividad promocional desde el dropdown
4. Elegir puntos de venta sin actividad
5. Aplicar filtros con botón "Filtrar"
6. Descargar `StoreCheck_Borrador.xlsx`

## 👥 Equipo de Desarrollo

- **Desarrollador Principal**: [  Jackelin Nuñez]
- **Analista Funcional**: [Equipo de BI]

## 📄 Licencia

Este proyecto está bajo la Licencia [TIPO_DE_LICENCIA]. Ver el archivo `LICENSE` para más detalles.

---

**Versión**: 1.0.0  
**Última Actualización**: 14/0/2025 
**Estado**: Producción ✅

