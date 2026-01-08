# 📋 Sistema de Entrega de Turno

Un sistema integral de **Google Apps Script** para gestionar entregas de turno, generar reportes en PDF profesionales y automatizar notificaciones por correo electrónico.

## 🎯 Características

✅ **Formulario Web Interactivo** - Interfaz moderna y responsiva para registrar entregas de turno  
✅ **Gestión de Vehículos** - Control y seguimiento de vehículos asignados a cada turno  
✅ **Generación de PDF** - Reportes profesionales con datos completos de la entrega  
✅ **Notificaciones Automáticas** - Envío de PDFs por correo electrónico a múltiples destinatarios  
✅ **Integración Google Sheets** - Almacenamiento automático de datos en hojas de cálculo  
✅ **Validación de Datos** - Controles de entrada para garantizar datos confiables  

## 📁 Estructura del Proyecto

```
├── CodigoGS.js          # Script principal (Google Apps Script)
├── PdfUtilsGS.js        # Utilidades para generación de PDFs
├── Formulario.html      # Interfaz del formulario web
├── pdfTemplate.html     # Plantilla HTML para PDFs
├── Style.css            # Estilos CSS del formulario
└── README.md            # Este archivo
```

## 📦 Archivos Principales

### **CodigoGS.js**
Script principal vinculado al Google Sheet. Contiene:
- Funciones de gestión del formulario (`doGet()`)
- Recuperación de registros desde Google Sheets
- Integración entre formulario y base de datos
- Configuración centralizada de correos y hojas

### **PdfUtilsGS.js**
Gestión completa de generación y envío de PDFs:
- Conversión de HTML a PDF profesional
- Generación de estadísticas de vehículos
- Envío automatizado de correos
- Plantillas optimizadas para impresión

### **Formulario.html**
Interfaz web responsiva que incluye:
- Campos para registro de datos de entrega
- Sección de vehículos con estados
- Áreas de observaciones y novedades
- Validación en tiempo real

### **pdfTemplate.html**
Plantilla HTML diseñada para conversión a PDF con:
- Diseño profesional y formal
- Secciones organizadas de información
- Estilos optimizados para impresión

### **Style.css**
Estilos CSS que definen:
- Colores corporativos
- Diseño responsivo
- Componentes de formulario personalizados

## 🚀 Cómo Usar

### 1. **Configuración Inicial**

#### Requisitos:
- Cuenta de Google con acceso a Google Apps Script
- Google Sheet vinculado al script
- Permisos para enviar correos desde la cuenta

#### Pasos:
1. **Crear un Google Sheet** con las siguientes hojas:
   - `Entregas_Turno` - Datos principales de entregas
   - `Entregas_Turno_Vehiculos` - Información de vehículos

2. **Vincular el script** al Google Sheet:
   - Ve a Extensiones → Apps Script
   - Copia el contenido de `CodigoGS.js` y `PdfUtilsGS.js`
   - Copia el contenido de `Formulario.html` y `pdfTemplate.html`

3. **Configurar correos** en `CodigoGS.js`:
   ```javascript
   const CONFIG = {
     HOJA_PRINCIPAL: "Entregas_Turno",
     HOJA_VEHICULOS: "Entregas_Turno_Vehiculos",
     CORREOS: ["correo1@example.com", "correo2@example.com"]
   };
   ```

### 2. **Usar el Formulario**

1. Ejecuta la función `doGet()` para obtener la URL web
2. Completa el formulario con:
   - **Facturador** - Nombre de la persona que entrega
   - **Fecha** - Día de la entrega
   - **Vehículos** - Lista de vehículos con estados
   - **Observaciones** - Notas adicionales
   - **Novedades** - Incidentes o cambios importantes

3. Haz clic en **"Generar PDF y Enviar"** para:
   - Guardar datos en Google Sheets
   - Generar PDF profesional
   - Enviar a todos los correos configurados

## 📊 Esquema de Datos

### Hoja: Entregas_Turno
| Campo | Tipo | Descripción |
|-------|------|-------------|
| Fecha | Date | Fecha de la entrega |
| Facturador | String | Nombre del responsable |
| Vehiculos | String | Lista de vehículos asignados |
| Observaciones | String | Notas generales |
| Novedades | String | Incidentes reportados |
| Timestamp | DateTime | Registro automático del envío |

### Hoja: Entregas_Turno_Vehiculos
| Campo | Tipo | Descripción |
|-------|------|-------------|
| Placa | String | Identificación del vehículo |
| Estado | String | Activo/Inactivo/Mantenimiento |
| KM | Number | Kilometraje actual |
| Observaciones | String | Notas del vehículo |

## 🔧 Funciones Principales

### En CodigoGS.js
- **`doGet()`** - Renderiza el formulario web
- **`getRegistros()`** - Obtiene todos los registros con vehículos
- **`guardarDatos(data)`** - Almacena datos en Google Sheets
- **`obtenerSpreadsheet()`** - Conexión segura a Google Sheet

### En PdfUtilsGS.js
- **`generateAndSendPDF(data)`** - Genera PDF y envía correos
- **`generateHTMLProfesionalCorregido(data)`** - Crea HTML optimizado
- **`sendEmailProfesional(data, pdfBlob)`** - Envía correos con PDF
- **`getVehiculosStatsObject(vehiculos)`** - Calcula estadísticas

## 📧 Notificaciones

El sistema envía automáticamente correos a:
- `soptransformaciondigital@pastascomarrico.com`
- `emorellanos@pastascomarrico.com`
- `eestevez@pastascomarrico.com`
- `mkjimenez@pastascomarrico.com`
- `jjguerrero@pastascomarrico.com`
- `practlogistica@pastascomarrico.com`
- `omorales@pastascomarrico.com`

*Nota: Modifica la lista de correos según tu necesidad en la configuración.*

## 🎨 Características de Diseño

- **Colores Corporativos**: Azul (#2c5aa0) y complementarios
- **Responsivo**: Funciona en desktop, tablet y móvil
- **Optimizado para PDF**: Estilos especiales para impresión
- **Accesible**: Estructura HTML semántica

## ⚙️ Configuración Avanzada

### Cambiar Fuentes
En `PdfUtilsGS.js`, modifica:
```javascript
font-family: 'Arial', 'Helvetica', sans-serif;
```

### Personalizar Estilos PDF
Edita los estilos CSS en la función `generateHTMLProfesionalCorregido()`

### Añadir Nuevos Campos
1. Agrega el input HTML en `Formulario.html`
2. Incluye la columna en Google Sheets
3. Actualiza la lógica de lectura/escritura en `CodigoGS.js`

## 🛠️ Troubleshooting

| Problema | Solución |
|----------|----------|
| "No se encontraron las hojas requeridas" | Verifica que las hojas se llamen exactamente `Entregas_Turno` y `Entregas_Turno_Vehiculos` |
| PDFs no se generan | Comprueba permisos de Google Apps Script |
| Correos no se envían | Verifica la lista de correos y conexión a internet |
| Datos no se guardan | Asegúrate que el script está vinculado al Sheet correcto |

## 📝 Notas de Versión

**v6.0-PROFESIONAL-CORREGIDO**
- ✅ HTML y PDF optimizados para impresión
- ✅ Fuentes legibles y profesionales
- ✅ Manejo robusto de errores
- ✅ Integración completa con Google Sheets

## 📄 Licencia

Desarrollado internamente para Pastas Comarrico.

## 👥 Contacto y Soporte

Para reportar problemas o solicitar mejoras, contacta a:
- **Email de Soporte**: siendokevi@gmail.com

## 👥 Creadores:

Creado por KEVIN CAMILO DELGADO RESTREPO

---

**Última actualización**: Enero 2026  
**Versión**: 6.0-PROFESIONAL-CORREGIDO
