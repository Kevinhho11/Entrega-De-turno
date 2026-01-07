/**
 * SISTEMA DE ENTREGA DE TURNO - VERSIÓN COMPLETA Y CORREGIDA
 * Script vinculado directamente al Google Sheet
 */

// ==================== CONFIGURACIÓN COMPLETA ====================
// ==================== CONFIGURACIÓN CORREGIDA ====================
const CONFIG = {
  HOJA_PRINCIPAL: "Entregas_Turno",
  HOJA_VEHICULOS: "Entregas_Turno_Vehiculos",  // ← CORREGIDO: con "s"
  CORREOS: ["soptransformaciondigital@pastascomarrico.com", "emorellanos@pastascomarrico.com", "eestevez@pastascomarrico.com", "mkjimenez@pastascomarrico.com", "jjguerrero@pastascomarrico.com", "practlogistica@pastascomarrico.com", "omorales@pastascomarrico.com"]
};

// ==================== FUNCIONES PRINCIPALES ====================

/**
 * Función principal - muestra el formulario
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Sistema de Entrega de Turno')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función para incluir archivos
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Función para guardar datos del formulario
 */
/**
 * Obtener el spreadsheet de forma segura - VERSIÓN CORREGIDA
 */
function obtenerSpreadsheet() {
  try {
    console.log('🔍 Buscando spreadsheet...');
    
    // Método 1: Spreadsheet activo (para scripts vinculados)
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (spreadsheet) {
      console.log('✅ Spreadsheet encontrado:', spreadsheet.getName());
      return spreadsheet;
    }
    
    throw new Error('No se pudo acceder al spreadsheet activo');
    
  } catch (error) {
    console.error('❌ Error obteniendo spreadsheet:', error);
    throw new Error('No se pudo conectar con Google Sheets: ' + error.toString());
  }
}

// === RESTAURAR getRegistros: Devuelve registros con vehículos y estados asociados, sin cambios de esquema ===
function getRegistros() {
  try {
    // Obtener spreadsheet y hojas
    const spreadsheet = obtenerSpreadsheet();
    const hojaTurnos = spreadsheet.getSheetByName(CONFIG.HOJA_PRINCIPAL);
    const hojaVehiculos = spreadsheet.getSheetByName(CONFIG.HOJA_VEHICULOS);
    if (!hojaTurnos || !hojaVehiculos) {
      return {
        success: false,
        error: "No se encontraron las hojas requeridas",
        data: []
      };
    }

    // Leer registros de turnos
    const datosTurnos = hojaTurnos.getDataRange().getValues();
    const headersTurnos = datosTurnos[0];
    const registros = datosTurnos.slice(1).map(row => {
      let obj = {};
      headersTurnos.forEach((h, i) => {
        let valor = row[i];
        if (valor instanceof Date) {
          valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (valor === null || valor === undefined) {
          valor = "";
        } else {
          valor = valor.toString();
        }
        obj[h] = valor;
      });
      return obj;
    });

    // Leer registros de vehículos
    const datosVehiculos = hojaVehiculos.getDataRange().getValues();
    const headersVehiculos = datosVehiculos[0];
    const vehiculos = datosVehiculos.slice(1).map(row => {
      let obj = {};
      headersVehiculos.forEach((h, i) => {
        let valor = row[i];
        if (valor instanceof Date) {
          valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (valor === null || valor === undefined) {
          valor = "";
        } else {
          valor = valor.toString();
        }
        obj[h] = valor;
      });
      return obj;
    });

    // Normalizar fecha a yyyy-MM-dd para comparar
    function normalizarSoloFecha(fechaStr) {
      if (!fechaStr) return '';
      // Si viene como 'yyyy-MM-dd HH:mm:ss' o 'yyyy-MM-dd'
      if (fechaStr.includes('-')) {
        const soloFecha = fechaStr.split(' ')[0];
        const partes = soloFecha.split('-');
        if (partes.length === 3) {
          return `${partes[0]}-${partes[1].padStart(2,'0')}-${partes[2].padStart(2,'0')}`;
        }
      }
      // Si viene como 'dd/MM/yyyy' o 'd/M/yyyy'
      if (fechaStr.includes('/')) {
        const soloFecha = fechaStr.split(' ')[0];
        const partes = soloFecha.split('/');
        if (partes.length === 3) {
          const [d, m, y] = partes;
          const anio = y.length === 2 ? '20'+y : y;
          return `${anio}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
        }
      }
      return fechaStr;
    }

    // Asociar vehículos a cada registro de turno por FECHA normalizada
    const registrosConVehiculos = registros.map(reg => {
      const fechaReg = normalizarSoloFecha(reg['FECHA']);
      const vehiculosAsociados = vehiculos.filter(v => normalizarSoloFecha(v['FECHA']) === fechaReg);
      // Contar estados
      const total = vehiculosAsociados.length;
      const confirmados = vehiculosAsociados.filter(v => (v['ESTADO']||'').toLowerCase() === 'confirmado').length;
      const pendientes = vehiculosAsociados.filter(v => (v['ESTADO']||'').toLowerCase() === 'pendiente').length;
      const cancelados = vehiculosAsociados.filter(v => (v['ESTADO']||'').toLowerCase() === 'cancelado').length;
      return {
        ...reg,
        vehiculos: vehiculosAsociados,
        vehiculosStats: {
          total,
          confirmados,
          pendientes,
          cancelados
        }
      };
    });

    return {
      success: true,
      data: registrosConVehiculos,
      total: registrosConVehiculos.length,
      message: `Se encontraron ${registrosConVehiculos.length} registros`
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      data: [],
      message: "Error crítico en el servidor"
    };
  }
}


/**
 * Obtener vehículos registrados
 */
function getVehiculosRegistros() {
  try {
    const spreadsheet = obtenerSpreadsheet();
    // USAR EL MISMO NOMBRE QUE EN CONFIG
    const hoja = spreadsheet.getSheetByName(CONFIG.HOJA_VEHICULOS);
    
    if (!hoja) {
      console.log('📭 Hoja de vehículos no encontrada:', CONFIG.HOJA_VEHICULOS);
      return { success: true, data: [] };
    }
    
    // Resto del código igual...
    // El resto del código igual...
    
    const datos = hoja.getDataRange().getValues();
    
    if (datos.length <= 1) {
      return { success: true, data: [] };
    }
    
    const encabezados = datos[0];
    const vehiculos = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const vehiculo = {};
      
      encabezados.forEach((header, index) => {
        vehiculo[header] = fila[index];
      });
      
      if (vehiculo['ID']) {
        vehiculos.push(vehiculo);
      }
    }
    
    return { success: true, data: vehiculos };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}


/**
 * Función de prueba de conexión
 */
function testConnection() {
  try {
    const spreadsheet = obtenerSpreadsheet();
    const hojas = spreadsheet.getSheets().map(s => s.getName());
    
    return {
      success: true,
      message: "✅ CONEXIÓN EXITOSA",
      spreadsheet: spreadsheet.getName(),
      hojas: hojas,
      config: CONFIG
    };
  } catch (error) {
    return {
      success: false,
      error: "❌ Error: " + error.toString()
    };
  }
}

// ...existing code...

// ==================== FUNCIONES AUXILIARES ====================

/**
 * Obtener el spreadsheet de forma segura
 */
/**
 * Obtener el spreadsheet de forma segura - VERSIÓN CORREGIDA
 */
function obtenerSpreadsheet() {
  try {
    console.log('🔍 Buscando spreadsheet...');
    
    // Método 1: Spreadsheet activo (para scripts vinculados)
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (spreadsheet) {
      console.log('✅ Spreadsheet encontrado:', spreadsheet.getName());
      return spreadsheet;
    }
    
    throw new Error('No se pudo acceder al spreadsheet activo');
    
  } catch (error) {
    console.error('❌ Error obteniendo spreadsheet:', error);
    throw new Error('No se pudo conectar con Google Sheets: ' + error.toString());
  }
}

/**
 * Guardar datos completos en ambas hojas
 */
/**
 * Guardar datos completos en ambas hojas - VERSIÓN CORREGIDA
 */
function guardarDatosCompletos(spreadsheet, datos) {
  // Asegurar que existan las hojas
  let hojaPrincipal = spreadsheet.getSheetByName(CONFIG.HOJA_PRINCIPAL);
  if (!hojaPrincipal) {
    hojaPrincipal = crearHojaPrincipal(spreadsheet);
  }
  
  let hojaVehiculos = spreadsheet.getSheetByName(CONFIG.HOJA_VEHICULOS);
  if (!hojaVehiculos) {
    hojaVehiculos = crearHojaVehiculos(spreadsheet);
  }
  
  const timestamp = new Date();
  const idEntrega = timestamp.getTime().toString();
  
  // CORRECCIÓN: Obtener fecha y hora ACTUALES, no de datos.fecha
  let fechaStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
  let horaStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");

  // CORRECCIÓN: Guardar en hoja principal con columnas en ORDEN CORRECTO
  const filaPrincipal = [
    timestamp,                          // Columna 1: HORA (timestamp completo)
    datos.facturador,                   // Columna 2: NOMBRE
    datos.turno,                        // Columna 3: TURNO
    datos.fecha,                        // Columna 4: FECHA (solo fecha del formulario)
    datos.observaciones,                // Columna 5: OBSERVACIONES ← AQUÍ ESTABA EL ERROR
    datos.facturasPendientes || "",     // Columna 6: FACTURAS FLETES PENDIENTES
    datos.novedadesMateriales || "",    // Columna 7: NOVEDADES MATERIALES
    (datos.puntosConsiderar || []).join(", "), // Columna 8: PUNTOS A CONSIDERAR
    datos.vehiculos.length,             // Columna 9: VEHICULOS (cantidad)
    datos.puntosConsiderar.length       // Columna 10: TOTAL_PUNTOS
  ];

  console.log('📝 Guardando en hoja principal:', filaPrincipal);
  
  hojaPrincipal.appendRow(filaPrincipal);
  const ultimaFila = hojaPrincipal.getLastRow();

  // Formatear fechas correctamente
  hojaPrincipal.getRange(ultimaFila, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  hojaPrincipal.getRange(ultimaFila, 4).setNumberFormat("dd/MM/yyyy");
  
  // CORRECCIÓN: Guardar vehículos con estructura correcta
  if (datos.vehiculos && datos.vehiculos.length > 0) {
    const datosVehiculos = datos.vehiculos.map(v => [
      idEntrega,                          // ID
      v.destino || "",                    // DESTINO
      v.transportadora || "",             // TRANSPORTADORA
      v.placa || "",                      // PLACA
      v.horaCitada || "",                 // HORA CITADA
      v.estado || "pendiente",            // ESTADO
      v.consecutivo || "",                // CONSECUTIVO
      datos.fecha,                        // FECHA (del formulario)
      horaStr                             // HORA_ENTREGA (hora actual)
    ]);
    
    hojaVehiculos.getRange(
      hojaVehiculos.getLastRow() + 1, 
      1, 
      datosVehiculos.length, 
      datosVehiculos[0].length
    ).setValues(datosVehiculos);
  }
  
  // Intentar enviar PDF
// En la función guardarDatosCompletos, cambia esta parte:
// Intentar enviar PDF
let pdfResult = { success: false, error: "No se intentó enviar" };
try {
  console.log('🔄 Intentando enviar PDF...');
  pdfResult = generateAndSendPDF(datos); // ← USA LA NUEVA FUNCIÓN
} catch (pdfError) {
  console.error('❌ Error enviando PDF:', pdfError);
  pdfResult = { 
    success: false, 
    error: pdfError.toString(),
    message: "Error al enviar PDF pero datos guardados"
  };
}
  
  return {
    success: true,
    id: idEntrega,
    message: "Datos guardados correctamente",
    pdfResult: pdfResult
  };
}

/**
 * Crear hoja principal con formato
 */
function crearHojaPrincipal(spreadsheet) {
  const hoja = spreadsheet.insertSheet(CONFIG.HOJA_PRINCIPAL);
  const encabezados = [
    "HORA", "NOMBRE", "TURNO", "FECHA", "HORA_ENTREGA", "OBSERVACIONES", 
    "FACTURAS FLETES PENDIENTES", "NOVEDADES MATERIALES", 
    "PUNTOS A CONSIDERAR", "TOTAL_VEHICULOS", "TOTAL_PUNTOS"
  ];
  
  const headerRange = hoja.getRange(1, 1, 1, encabezados.length);
  headerRange.setValues([encabezados]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e40af");
  headerRange.setFontColor("white");
  
  return hoja;
}

/**
 * Crear hoja de vehículos con formato
 */
function crearHojaVehiculos(spreadsheet) {
  const hoja = spreadsheet.insertSheet(CONFIG.HOJA_VEHICULOS);
  const encabezados = [
    "ID", "DESTINO", "TRANSPORTADORA", "PLACA", 
    "HORA CITADA", "ESTADO", "CONSECUTIVO", "FECHA", "HORA_ENTREGA"
  ];
  
  const headerRange = hoja.getRange(1, 1, 1, encabezados.length);
  headerRange.setValues([encabezados]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e40af");
  headerRange.setFontColor("white");
  
  return hoja;
}

/**
 * Enviar PDF por correo
 */
/**
 * Enviar PDF por correo - VERSIÓN PROFESIONAL CON DISEÑO
 */
/**
 * Enviar PDF por correo - VERSIÓN QUE USA EL NUEVO DISEÑO
 */
function enviarPDF(datos) {
  try {
    console.log('🚀 USANDO NUEVO DISEÑO DE PDF DESDE FRONTEND');
    
    // Usar directamente el nuevo método que SÍ funciona
    const result = generateAndSendPDF(datos);
    
    console.log('✅ Resultado del PDF nuevo:', result);
    return result;
    
  } catch (error) {
    console.error('❌ Error en enviarPDF:', error);
    
    // Fallback: intentar con el método viejo si el nuevo falla
    try {
      console.log('🔄 Intentando método de respaldo...');
      return generarYEnviarPDFRespaldo(datos);
    } catch (fallbackError) {
      return {
        success: false,
        error: `Nuevo: ${error.toString()}, Respaldo: ${fallbackError.toString()}`
      };
    }
  }
}

/**
 * Método de respaldo por si falla el nuevo
 */
function generarYEnviarPDFRespaldo(datos) {
  // Tu código viejo de enviarPDF aquí como respaldo
  try {
    console.log('🔄 Usando método de respaldo...');
    
    const htmlContent = generarHTMLProfesional(datos);
    const pdf = Utilities.newBlob(htmlContent, MimeType.HTML, 'entrega_turno.pdf')
                      .getAs(MimeType.PDF);
    
    let emailsEnviados = 0;
    CONFIG.CORREOS.forEach(email => {
      try {
        const emailResult = enviarCorreoProfesional(email, datos, pdf);
        if (emailResult.success) {
          emailsEnviados++;
        }
      } catch (e) {
        console.log('Error enviando a', email, e);
      }
    });
    
    return {
      success: true,
      emailsSent: emailsEnviados,
      message: `PDF (respaldo) enviado a ${emailsEnviados} destinatarios`
    };
    
  } catch (error) {
    console.error('❌ Error en respaldo:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}









/**
 * GENERAR Y ENVIAR PDF - FUNCIÓN PRINCIPAL CORREGIDA
 */
function generateAndSendPDF(datos) {
  try {
    console.log('📧 INICIANDO ENVÍO DE PDF...', datos);
    
    // Validar datos esenciales
    if (!datos.facturador || !datos.turno || !datos.fecha) {
      throw new Error('Datos incompletos para generar PDF');
    }
    
    // 1. GENERAR EL HTML PARA EL PDF
    console.log('🔄 Generando HTML del PDF...');
    const htmlContent = generarHTMLProfesional(datos);
    
    if (!htmlContent) {
      throw new Error('No se pudo generar el contenido HTML');
    }
    
    // 2. CREAR EL BLOB PDF
    console.log('📄 Creando archivo PDF...');
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML, 'temp.html');
    const pdf = blob.getAs(MimeType.PDF);
    pdf.setName(`Entrega_Turno_${datos.facturador}_${datos.fecha.replace(/-/g, '')}.pdf`);
    
    // 3. ENVIAR POR CORREO
    console.log('📤 Enviando correos a:', CONFIG.CORREOS);
    
    let emailsEnviados = 0;
    let errores = [];
    
    CONFIG.CORREOS.forEach(email => {
      try {
        console.log(`📨 Enviando a: ${email}`);
        const resultado = enviarCorreoProfesional(email, datos, pdf);
        
        if (resultado && resultado.success) {
          emailsEnviados++;
          console.log(`✅ Correo enviado a: ${email}`);
        } else {
          errores.push(`${email}: Falló`);
        }
      } catch (emailError) {
        console.error(`❌ Error enviando a ${email}:`, emailError);
        errores.push(`${email}: ${emailError.toString()}`);
      }
    });
    
    // 4. VERIFICAR RESULTADO
    if (emailsEnviados === 0) {
      throw new Error(`No se pudo enviar a ningún destinatario. Errores: ${errores.join(', ')}`);
    }
    
    console.log(`✅ ENVÍO COMPLETADO: ${emailsEnviados}/${CONFIG.CORREOS.length} correos enviados`);
    
    return {
      success: true,
      emailsSent: emailsEnviados,
      totalRecipients: CONFIG.CORREOS.length,
      errors: errores,
      message: `PDF enviado exitosamente a ${emailsEnviados} destinatarios`
    };
    
  } catch (error) {
    console.error('❌ ERROR CRÍTICO EN generateAndSendPDF:', error);
    
    // Intentar método alternativo
    try {
      console.log('🔄 Intentando método alternativo...');
      return generarYEnviarPDFRespaldo(datos);
    } catch (fallbackError) {
      console.error('❌ También falló el método alternativo:', fallbackError);
      
      return {
        success: false,
        error: `Error principal: ${error.toString()}. Fallback: ${fallbackError.toString()}`,
        emailsSent: 0
      };
    }
  }
}









/**
 * VERIFICAR ESTADO DE ENVÍO DE CORREOS
 */
function verificarEstadoCorreos() {
  try {
    console.log('🔍 VERIFICANDO ESTADO DE CORREOS...');
    
    // Verificar límites de cuota
    const emailQuota = MailApp.getRemainingDailyQuota();
    console.log(`📧 Cuota restante de correos: ${emailQuota}`);
    
    if (emailQuota < CONFIG.CORREOS.length) {
      return {
        success: false,
        error: `Cuota insuficiente. Necesitas ${CONFIG.CORREOS.length}, tienes ${emailQuota}`,
        quota: emailQuota,
        required: CONFIG.CORREOS.length
      };
    }
    
    // Verificar permisos
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "Prueba de permisos - Sistema Entrega Turno",
        body: "Este es un correo de prueba para verificar permisos."
      });
    } catch (permisoError) {
      return {
        success: false,
        error: "Problema de permisos: " + permisoError.toString()
      };
    }
    
    return {
      success: true,
      quota: emailQuota,
      required: CONFIG.CORREOS.length,
      message: `✅ Estado OK - Cuota suficiente: ${emailQuota} correos disponibles`
    };
    
  } catch (error) {
    return {
      success: false,
      error: "Error verificando estado: " + error.toString()
    };
  }
}



/**
 * Generar HTML profesional para el PDF
 */
/**
 * Generar HTML optimizado para PDF
 */
function generarHTMLProfesional(datos) {
  const fechaFormateada = new Date(datos.fecha).toLocaleDateString('es-ES', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  
  const stats = obtenerEstadisticasVehiculos(datos.vehiculos);
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      line-height: 1.6;
      color: #1f2937;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      padding: 20px;
      min-height: 100vh;
    }
    
    .container {
      max-width: 1000px;
      margin: 0 auto;
      background: white;
      border-radius: 20px;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
      overflow: hidden;
      position: relative;
    }
    
    /* Header con gradiente moderno */
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 40px;
      text-align: center;
      position: relative;
      overflow: hidden;
    }
    
    .header::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 100" fill="rgba(255,255,255,0.1)"><polygon points="0,0 1000,50 1000,100 0,100"/></svg>');
      background-size: cover;
    }
    
    .header h1 {
      font-size: 2.5em;
      font-weight: 700;
      margin-bottom: 10px;
      position: relative;
      text-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    
    .header p {
      font-size: 1.1em;
      opacity: 0.9;
      font-weight: 300;
      position: relative;
    }
    
    /* Contenido principal */
    .content {
      padding: 40px;
    }
    
    /* Secciones con diseño de tarjetas */
    .section {
      background: white;
      border-radius: 16px;
      padding: 30px;
      margin-bottom: 30px;
      border: 1px solid #e5e7eb;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
      transition: all 0.3s ease;
      position: relative;
    }
    
    .section:hover {
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
      transform: translateY(-2px);
    }
    
    .section-header {
      display: flex;
      align-items: center;
      margin-bottom: 25px;
      padding-bottom: 15px;
      border-bottom: 2px solid #f1f5f9;
    }
    
    .section-icon {
      font-size: 2em;
      margin-right: 15px;
      background: linear-gradient(135deg, #667eea, #764ba2);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }
    
    .section-title {
      font-size: 1.5em;
      font-weight: 600;
      color: #1e293b;
      margin: 0;
    }
    
    /* Grid de información */
    .info-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 20px;
      margin-top: 20px;
    }
    
    .info-card {
      background: linear-gradient(135deg, #f8fafc, #f1f5f9);
      padding: 20px;
      border-radius: 12px;
      border-left: 4px solid #667eea;
      transition: all 0.3s ease;
    }
    
    .info-card:hover {
      transform: translateX(5px);
      background: linear-gradient(135deg, #f1f5f9, #e2e8f0);
    }
    
    .info-label {
      font-size: 0.85em;
      color: #64748b;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 5px;
    }
    
    .info-value {
      font-size: 1.2em;
      font-weight: 600;
      color: #1e293b;
    }
    
    /* Estadísticas con diseño moderno */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
      gap: 15px;
      margin: 25px 0;
    }
    
    .stat-card {
      background: linear-gradient(135deg, #667eea, #764ba2);
      color: white;
      padding: 20px;
      border-radius: 12px;
      text-align: center;
      box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
      transition: all 0.3s ease;
    }
    
    .stat-card:hover {
      transform: translateY(-3px);
      box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }
    
    .stat-number {
      font-size: 2em;
      font-weight: 700;
      margin-bottom: 5px;
    }
    
    .stat-label {
      font-size: 0.85em;
      opacity: 0.9;
      font-weight: 500;
    }
    
    /* Tablas mejoradas */
    .table-container {
      overflow-x: auto;
      margin-top: 20px;
      border-radius: 12px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .vehiculos-table {
      width: 100%;
      border-collapse: collapse;
      background: white;
      border-radius: 12px;
      overflow: hidden;
    }
    
    .vehiculos-table thead {
      background: linear-gradient(135deg, #667eea, #764ba2);
      color: white;
    }
    
    .vehiculos-table th {
      padding: 15px 12px;
      text-align: left;
      font-weight: 600;
      font-size: 0.9em;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .vehiculos-table td {
      padding: 12px;
      border-bottom: 1px solid #f1f5f9;
      font-size: 0.9em;
    }
    
    .vehiculos-table tbody tr {
      transition: all 0.3s ease;
    }
    
    .vehiculos-table tbody tr:hover {
      background: #f8fafc;
      transform: scale(1.01);
    }
    
    .vehiculos-table tbody tr:nth-child(even) {
      background: #fafafa;
    }
    
    .vehiculos-table tbody tr:nth-child(even):hover {
      background: #f1f5f9;
    }
    
    /* Badges de estado mejorados */
    .status-badge {
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 0.8em;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .status-confirmado {
      background: linear-gradient(135deg, #10b981, #059669);
      color: white;
    }
    
    .status-pendiente {
      background: linear-gradient(135deg, #f59e0b, #d97706);
      color: white;
    }
    
    .status-cancelado {
      background: linear-gradient(135deg, #ef4444, #dc2626);
      color: white;
    }
    
    /* Lista de puntos mejorada */
    .puntos-list {
      list-style: none;
      margin-top: 15px;
    }
    
    .punto-item {
      display: flex;
      align-items: flex-start;
      padding: 15px;
      margin-bottom: 10px;
      background: #f8fafc;
      border-radius: 10px;
      border-left: 4px solid #10b981;
      transition: all 0.3s ease;
    }
    
    .punto-item:hover {
      background: #f1f5f9;
      transform: translateX(5px);
    }
    
    .punto-icon {
      font-size: 1.2em;
      margin-right: 15px;
      margin-top: 2px;
      color: #10b981;
    }
    
    .punto-text {
      flex: 1;
      color: #374151;
      font-weight: 500;
    }
    
    /* Contenido de texto */
    .text-content {
      background: #f8fafc;
      padding: 20px;
      border-radius: 10px;
      border-left: 4px solid #667eea;
      line-height: 1.7;
      color: #374151;
    }
    
    /* Footer mejorado */
    .footer {
      background: linear-gradient(135deg, #1e293b, #0f172a);
      color: white;
      padding: 30px 40px;
      text-align: center;
    }
    
    .footer p {
      margin-bottom: 10px;
      opacity: 0.8;
    }
    
    .footer-logo {
      font-size: 1.2em;
      font-weight: 700;
      color: #667eea;
      letter-spacing: 1px;
    }
    
    /* Utilidades */
    .text-center {
      text-align: center;
    }
    
    .keep-together {
      page-break-inside: avoid;
    }
    
    /* Columnas específicas para tabla */
    .col-num { width: 50px; }
    .col-destino { width: 150px; }
    .col-transportadora { width: 150px; }
    .col-placa { width: 100px; }
    .col-hora { width: 100px; }
    .col-estado { width: 120px; }
    .col-consecutivo { width: 80px; }
    
    /* Efectos de impresión */
    @media print {
      body {
        background: white !important;
        padding: 0 !important;
      }
      .container {
        box-shadow: none !important;
        border-radius: 0 !important;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header">
      <h1>📋 ENTREGA DE TURNO</h1>
      <p>Reporte Oficial - Sistema de Gestión Logística</p>
    </div>
    
    <!-- Content -->
    <div class="content">
      <!-- Información Básica -->
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">👤</div>
          <h2 class="section-title">Información Básica</h2>
        </div>
        <div class="info-grid">
          <div class="info-card">
            <div class="info-label">Facturador</div>
            <div class="info-value">${datos.facturador}</div>
          </div>
          <div class="info-card">
            <div class="info-label">Turno</div>
            <div class="info-value">${datos.turno.toUpperCase()}</div>
          </div>
          <div class="info-card">
            <div class="info-label">Fecha</div>
            <div class="info-value">${fechaFormateada.split(',')[0]}</div>
          </div>
          <div class="info-card">
            <div class="info-label">Hora Reporte</div>
            <div class="info-value">${new Date().toLocaleTimeString('es-ES', {hour: '2-digit', minute:'2-digit'})}</div>
          </div>
        </div>
      </div>

      <!-- Observaciones -->
      ${datos.observaciones ? `
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">📝</div>
          <h2 class="section-title">Observaciones</h2>
        </div>
        <div class="text-content">
          ${datos.observaciones}
        </div>
      </div>
      ` : ''}

      <!-- Vehículos Cargados Durante el Turno -->
      ${datos.vehiculosCargados && datos.vehiculosCargados.length > 0 ? `
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">✅</div>
          <h2 class="section-title">Vehículos Cargados Durante el Turno</h2>
        </div>
        <div class="table-container">
          <table class="vehiculos-table">
            <thead>
              <tr>
                <th class="col-num">#</th>
                <th class="col-destino">Destino</th>
                <th class="col-transportadora">Transportadora</th>
                <th class="col-placa">Placa</th>
                <th class="col-hora">Quién Cargó</th>
                <th class="col-estado">Quién Preparó</th>
                <th class="col-consecutivo">Estatus</th>
              </tr>
            </thead>
            <tbody>
              ${generarFilasVehiculosCargados(datos.vehiculosCargados)}
            </tbody>
          </table>
        </div>
      </div>
      ` : ''}

      <!-- Vehículos (Proyección Día Siguiente) -->
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">🚚</div>
          <h2 class="section-title">Proyección de Vehículos</h2>
        </div>
        
        <!-- Estadísticas -->
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-number">${stats.total}</div>
            <div class="stat-label">Total</div>
          </div>
          <div class="stat-card">
            <div class="stat-number">${stats.confirmados}</div>
            <div class="stat-label">Confirmados</div>
          </div>
          <div class="stat-card">
            <div class="stat-number">${stats.pendientes}</div>
            <div class="stat-label">Pendientes</div>
          </div>
          <div class="stat-card">
            <div class="stat-number">${stats.cancelados}</div>
            <div class="stat-label">Cancelados</div>
          </div>
        </div>
        
        <!-- Tabla de Vehículos -->
        <div class="table-container">
          <table class="vehiculos-table">
            <thead>
              <tr>
                <th class="col-num">#</th>
                <th class="col-destino">Destino</th>
                <th class="col-transportadora">Transportadora</th>
                <th class="col-placa">Placa</th>
                <th class="col-hora">Hora</th>
                <th class="col-estado">Estado</th>
                <th class="col-consecutivo">Cons.</th>
              </tr>
            </thead>
            <tbody>
              ${generarFilasVehiculosOptimizada(datos.vehiculos)}
            </tbody>
          </table>
        </div>
      </div>

      <!-- Facturas Pendientes -->
      ${datos.facturasPendientes ? `
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">📄</div>
          <h2 class="section-title">Facturas Pendientes</h2>
        </div>
        <div class="text-content">
          ${datos.facturasPendientes}
        </div>
      </div>
      ` : ''}

      <!-- Novedades Materiales -->
      ${datos.novedadesMateriales ? `
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">⚠️</div>
          <h2 class="section-title">Novedades Materiales</h2>
        </div>
        <div class="text-content">
          ${datos.novedadesMateriales}
        </div>
      </div>
      ` : ''}

      <!-- Puntos Completados -->
      ${datos.puntosConsiderar && datos.puntosConsiderar.length > 0 ? `
      <div class="section keep-together">
        <div class="section-header">
          <div class="section-icon">✅</div>
          <h2 class="section-title">Puntos Completados</h2>
        </div>
        <ul class="puntos-list">
          ${generarPuntosCompletadosOptimizada(datos.puntosConsiderar)}
        </ul>
      </div>
      ` : ''}
    </div>

    <!-- Footer -->
    <div class="footer">
      <p>Reporte generado automáticamente - Sistema de Entrega de Turno</p>
      <div class="footer-logo">COMARRICO LOGÍSTICA</div>
    </div>
  </div>
</body>
</html>
  `;
}
/**
 * Generar filas de vehículos cargados durante el turno
 */
function generarFilasVehiculosCargados(vehiculosCargados) {
  if (!vehiculosCargados || vehiculosCargados.length === 0) {
    return '<tr><td colspan="7" class="text-center" style="padding: 20px; color: #64748b;">No hay vehículos cargados registrados</td></tr>';
  }
  return vehiculosCargados.map((vehiculo, index) => {
    return `
      <tr>
        <td class="col-num"><strong>${index + 1}</strong></td>
        <td class="col-destino">${vehiculo.destino || '-'}</td>
        <td class="col-transportadora">${vehiculo.transportadora || '-'}</td>
        <td class="col-placa">${vehiculo.placa || '-'}</td>
        <td class="col-hora">${vehiculo.quienCargo || '-'}</td>
        <td class="col-estado">${vehiculo.quienPreparo || '-'}</td>
        <td class="col-consecutivo">${vehiculo.estatus || vehiculo.estado || '-'}</td>
      </tr>
    `;
  }).join('');
}


/**
 * Generar filas de vehículos optimizada para PDF
 */
function generarFilasVehiculosOptimizada(vehiculos) {
  if (!vehiculos || vehiculos.length === 0) {
    return '<tr><td colspan="7" class="text-center" style="padding: 20px; color: #64748b;">No hay vehículos registrados</td></tr>';
  }
  
  return vehiculos.map((vehiculo, index) => {
    const estadoClass = `status-${vehiculo.estado || 'pendiente'}`;
    const estadoText = getEstadoTextoOptimizado(vehiculo.estado);
    
    // Acortar textos largos para que quepan
    const destino = vehiculo.destino && vehiculo.destino.length > 15 ? 
                   vehiculo.destino.substring(0, 15) + '...' : vehiculo.destino || '-';
    
    const transportadora = vehiculo.transportadora && vehiculo.transportadora.length > 15 ? 
                          vehiculo.transportadora.substring(0, 15) + '...' : vehiculo.transportadora || '-';
    
    return `
      <tr>
        <td class="col-num"><strong>${index + 1}</strong></td>
        <td class="col-destino" title="${vehiculo.destino || ''}">${destino}</td>
        <td class="col-transportadora" title="${vehiculo.transportadora || ''}">${transportadora}</td>
        <td class="col-placa"><code>${vehiculo.placa || '-'}</code></td>
        <td class="col-hora">${vehiculo.horaCitada || '-'}</td>
        <td class="col-estado"><span class="status-badge ${estadoClass}">${estadoText}</span></td>
        <td class="col-consecutivo">${vehiculo.consecutivo || '-'}</td>
      </tr>
    `;
  }).join('');
}

/**
 * Generar puntos completados optimizada
 */
function generarPuntosCompletadosOptimizada(puntos) {
  if (!puntos || puntos.length === 0) {
    return '<li class="punto-item"><div class="punto-icon">ℹ️</div><div class="punto-text">No se completaron actividades específicas</div></li>';
  }
  
  // Limitar a 5 puntos máximo para evitar overflow
  const puntosMostrar = puntos.slice(0, 5);
  
  return puntosMostrar.map(punto => {
    // Acortar texto si es muy largo
    const texto = punto.length > 60 ? punto.substring(0, 60) + '...' : punto;
    return `
      <li class="punto-item">
        <div class="punto-icon">✓</div>
        <div class="punto-text">${texto}</div>
      </li>
    `;
  }).join('');
}

/**
 * Obtener texto del estado optimizado
 */
function getEstadoTextoOptimizado(estado) {
  const estados = {
    'confirmado': '✓',
    'pendiente': '⏳', 
    'cancelado': '✗'
  };
  return estados[estado] || '⏳';
}

/**
 * Generar filas de la tabla de vehículos
 */
function generarFilasVehiculos(vehiculos) {
  if (!vehiculos || vehiculos.length === 0) {
    return '<tr><td colspan="7" style="text-align: center; color: #64748b; padding: 30px;">No hay vehículos registrados</td></tr>';
  }
  
  return vehiculos.map((vehiculo, index) => {
    const estadoClass = `status-${vehiculo.estado || 'pendiente'}`;
    const estadoText = getEstadoText(vehiculo.estado);
    
    return `
      <tr>
        <td><strong>${index + 1}</strong></td>
        <td>${vehiculo.destino || '-'}</td>
        <td>${vehiculo.transportadora || '-'}</td>
        <td><code>${vehiculo.placa || '-'}</code></td>
        <td>${vehiculo.horaCitada || '-'}</td>
        <td><span class="status-badge ${estadoClass}">${estadoText}</span></td>
        <td>${vehiculo.consecutivo || '-'}</td>
      </tr>
    `;
  }).join('');
}

/**
 * Generar puntos completados
 */
function generarPuntosCompletados(puntos) {
  if (!puntos || puntos.length === 0) {
    return '<div class="punto-item"><div class="punto-icon">ℹ️</div><div class="punto-text">No se completaron actividades específicas en este turno</div></div>';
  }
  
  return puntos.map(punto => `
    <div class="punto-item">
      <div class="punto-icon">✅</div>
      <div class="punto-text">${punto}</div>
    </div>
  `).join('');
}

/**
 * Obtener texto del estado
 */
function getEstadoText(estado) {
  const estados = {
    'confirmado': '✅ CONFIRMADO',
    'pendiente': '⏳ PENDIENTE', 
    'cancelado': '❌ CANCELADO'
  };
  return estados[estado] || '⏳ PENDIENTE';
}

/**
 * Obtener estadísticas de vehículos
 */
function obtenerEstadisticasVehiculos(vehiculos) {
  return {
    total: vehiculos.length,
    confirmados: vehiculos.filter(v => v.estado === 'confirmado').length,
    pendientes: vehiculos.filter(v => v.estado === 'pendiente').length,
    cancelados: vehiculos.filter(v => v.estado === 'cancelado').length
  };
}

/**
 * Enviar correo profesional
 */
function enviarCorreoProfesional(email, datos, pdf) {
  const subject = `🚚 Entrega de Turno - ${datos.facturador} - ${datos.turno.toUpperCase()}`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            line-height: 1.6;
            color: #1f2937;
            background: #f8fafc;
            margin: 0;
            padding: 0;
        }
        
        .container {
            max-width: 600px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
        }
        
        .header {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 28px;
            font-weight: 700;
            margin: 0;
        }
        
        .content {
            padding: 40px;
        }
        
        .info-card {
            background: #f8fafc;
            border-radius: 12px;
            padding: 25px;
            margin: 20px 0;
            border: 1px solid #e2e8f0;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 15px;
            margin: 20px 0;
        }
        
        .stat-item {
            text-align: center;
            padding: 15px;
            background: white;
            border-radius: 10px;
            border: 1px solid #e2e8f0;
        }
        
        .stat-number {
            font-size: 24px;
            font-weight: 700;
            color: #2563eb;
        }
        
        .stat-label {
            font-size: 12px;
            color: #64748b;
            text-transform: uppercase;
            font-weight: 600;
            margin-top: 5px;
        }
        
        .footer {
            background: #1e293b;
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .btn {
            display: inline-block;
            background: linear-gradient(135deg, #2563eb, #3b82f6);
            color: white;
            padding: 12px 30px;
            text-decoration: none;
            border-radius: 8px;
            font-weight: 600;
            margin: 20px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 NUEVA ENTREGA DE TURNO</h1>
            <p>Sistema de Gestión Logística</p>
        </div>
        
        <div class="content">
            <h2 style="color: #1e293b; margin-bottom: 20px;">¡Hola! Se ha registrado una nueva entrega de turno</h2>
            
            <div class="info-card">
                <h3 style="color: #2563eb; margin-bottom: 15px;">📊 Resumen del Reporte</h3>
                
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-number">${datos.facturador}</div>
                        <div class="stat-label">Facturador</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${datos.turno.toUpperCase()}</div>
                        <div class="stat-label">Turno</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${datos.vehiculos.length}</div>
                        <div class="stat-label">Vehículos</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${datos.puntosConsiderar.length}</div>
                        <div class="stat-label">Puntos Completados</div>
                    </div>
                </div>
            </div>
            
            <p style="color: #475569; margin: 20px 0;">
                Se adjunta el reporte completo en formato PDF con todos los detalles de la entrega de turno, 
                incluyendo la proyección de vehículos, observaciones y actividades completadas.
            </p>
            
            <div style="background: #dcfce7; padding: 20px; border-radius: 12px; border-left: 4px solid #10b981;">
                <p style="margin: 0; color: #166534; font-weight: 500;">
                    📎 <strong>Archivo adjunto:</strong> Entrega_Turno_${datos.facturador}_${datos.fecha.replace(/-/g, '')}.pdf
                </p>
            </div>
        </div>
        
        <div class="footer">
            <p style="margin: 0; opacity: 0.8;">Este correo fue generado automáticamente por el Sistema de Entrega de Turno</p>
            <p style="margin: 10px 0 0 0; opacity: 0.6; font-size: 12px;">
                Generado el ${new Date().toLocaleString('es-ES')}
            </p>
        </div>
    </div>
</body>
</html>
  `;
  
  const textBody = `
NUEVA ENTREGA DE TURNO REGISTRADA

📊 RESUMEN:
- Facturador: ${datos.facturador}
- Turno: ${datos.turno.toUpperCase()} 
- Fecha: ${new Date(datos.fecha).toLocaleDateString('es-ES')}
- Vehículos registrados: ${datos.vehiculos.length}
- Puntos completados: ${datos.puntosConsiderar.length}

Se adjunta el reporte completo en formato PDF con todos los detalles.

Este es un mensaje automático del Sistema de Entrega de Turno.
Generado el: ${new Date().toLocaleString('es-ES')}
  `;
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: textBody,
    htmlBody: htmlBody,
    attachments: [pdf],
    name: 'Sistema de Entrega de Turno'
  });
  
  return { success: true };
}
function diagnosticarProblema() {
  try {
    console.log('🔍 INICIANDO DIAGNÓSTICO COMPLETO...');
    
    // 1. Verificar acceso al Sheet
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      console.log('✅ Sheet abierto:', spreadsheet.getName());
    } catch (sheetError) {
      console.log('❌ Error abriendo Sheet:', sheetError.toString());
      return {
        success: false,
        error: 'No se puede acceder al Sheet: ' + sheetError.toString()
      };
    }
    
    // 2. Verificar hojas
    const sheet = spreadsheet.getSheetByName("Entregas_Turno");
    const vehiculosSheet = spreadsheet.getSheetByName("Entregas_Turno_Vehiculos");
    
    console.log('📊 Hoja principal existe:', !!sheet);
    console.log('🚚 Hoja vehículos existe:', !!vehiculosSheet);
    
    // 3. Verificar datos
    let datosPrincipal = [];
    let datosVehiculos = [];
    
    if (sheet) {
      datosPrincipal = sheet.getDataRange().getValues();
      console.log('📈 Datos principal:', datosPrincipal.length, 'filas');
    }
    
    if (vehiculosSheet) {
      datosVehiculos = vehiculosSheet.getDataRange().getValues();
      console.log('📈 Datos vehículos:', datosVehiculos.length, 'filas');
    }
    
    return {
      success: true,
      diagnostico: {
        spreadsheet: spreadsheet.getName(),
        hojas: {
          principal: {
            existe: !!sheet,
            filas: datosPrincipal.length,
            encabezados: datosPrincipal[0] || []
          },
          vehiculos: {
            existe: !!vehiculosSheet,
            filas: datosVehiculos.length,
            encabezados: datosVehiculos[0] || []
          }
        }
      }
    };
    
  } catch (error) {
    console.error('❌ Error en diagnóstico:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * GUARDAR ENTREGA DE TURNO - Función principal que llama el HTML
 */
function guardarEntregaTurno(datos) {
  console.log('📝 Recibiendo datos para guardar...', datos);
  
  try {
    // Validar datos básicos
    if (!datos.facturador || !datos.turno || !datos.fecha) {
      throw new Error('Faltan datos obligatorios: facturador, turno o fecha');
    }
    
    // Obtener spreadsheet
    const spreadsheet = obtenerSpreadsheet();
    console.log('✅ Spreadsheet obtenido:', spreadsheet.getName());
    
    // Guardar datos completos
    const resultado = guardarDatosCompletos(spreadsheet, datos);
    console.log('✅ Datos guardados:', resultado);
    
    return resultado;
    
  } catch (error) {
    console.error('❌ Error en guardarEntregaTurno:', error);
    return {
      success: false,
      error: error.toString(),
      message: "Error al guardar los datos"
    };
  }
}


/**
 * PRUEBA DE COMUNICACIÓN FRONTEND
 */
function testFrontendCommunication() {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: "✅ Comunicación exitosa con frontend",
    timestamp: new Date().toISOString(),
    testData: [
      { id: 1, nombre: "Test 1", fecha: new Date().toISOString() },
      { id: 2, nombre: "Test 2", fecha: new Date().toISOString() }
    ]
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * FUNCIÓN DE RESPUESTA SIMPLE - GARANTIZADA
 */
function getRegistrosSimple() {
  try {
    console.log('🎯 GET_REGISTROS_SIMPLE llamado');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Entregas_Turno");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    const headers = data[0];
    const registros = [];
    
    // Procesar todos los registros (excepto encabezado)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const registro = {};
      for (let j = 0; j < headers.length; j++) {
        registro[headers[j]] = row[j] ? row[j].toString() : "";
      }
      registros.push(registro);
    }
    
    console.log('✅ getRegistrosSimple retornando:', registros.length, 'registros');
    
    return {
      success: true,
      data: registros,
      message: "Datos de prueba cargados"
    };
    
  } catch (error) {
    console.error('❌ Error en getRegistrosSimple:', error);
    return {
      success: false,
      error: error.toString(),
      data: []
    };
  }
}

/**
 * DIAGNÓSTICO COMPLETO DEL SISTEMA
 */
function diagnosticoCompleto() {
  console.log('🔍 INICIANDO DIAGNÓSTICO COMPLETO...');
  
  const resultados = {
    spreadsheet: null,
    hojas: {},
    correos: {},
    permisos: {}
  };
  
  try {
    // 1. Verificar Spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    resultados.spreadsheet = {
      nombre: ss.getName(),
      id: ss.getId(),
      url: ss.getUrl()
    };
    console.log('✅ Spreadsheet:', resultados.spreadsheet);
    
    // 2. Verificar Hojas
    const hojas = ['Entregas_Turno', 'Entregas_Turno_Vehiculos'];
    hojas.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      resultados.hojas[nombreHoja] = {
        existe: !!hoja,
        filas: hoja ? hoja.getLastRow() : 0,
        columnas: hoja ? hoja.getLastColumn() : 0
      };
    });
    console.log('✅ Hojas:', resultados.hojas);
    
    // 3. Verificar Correos
    resultados.correos = verificarEstadoCorreos();
    console.log('✅ Estado correos:', resultados.correos);
    
    // 4. Verificar función generateAndSendPDF
    try {
      // Test con datos mínimos
      const testData = {
        facturador: "TEST",
        turno: "diurno", 
        fecha: "2024-01-01",
        vehiculos: [],
        puntosConsiderar: [],
        observaciones: "Prueba de diagnóstico"
      };
      
      // Solo verificar que la función existe y puede ser llamada
      if (typeof generateAndSendPDF === 'function') {
        resultados.pdfFunction = "✅ Función disponible";
      } else {
        resultados.pdfFunction = "❌ Función NO disponible";
      }
    } catch (funcError) {
      resultados.pdfFunction = "❌ Error: " + funcError.toString();
    }
    
    console.log('✅ Diagnóstico completado');
    return {
      success: true,
      diagnostico: resultados,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ Error en diagnóstico:', error);
    return {
      success: false,
      error: error.toString(),
      diagnostico: resultados
    };
  }
}