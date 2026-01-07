/**
 * UTILIDADES PARA GENERACIÓN DE PDF - VERSIÓN PROFESIONAL CORREGIDA
 * Diseño optimizado para PDF con fuentes legibles y estilo profesional
 */

const PDF_TEMPLATE_VERSION = '6.0-PROFESIONAL-CORREGIDO';
const CORREOS_NOTIFICACION = ["soptransformaciondigital@pastascomarrico.com", "emorellanos@pastascomarrico.com", "eestevez@pastascomarrico.com", "mkjimenez@pastascomarrico.com", "jjguerrero@pastascomarrico.com", "practlogistica@pastascomarrico.com", "omorales@pastascomarrico.com"]

/**
 * Genera un PDF PROFESIONAL optimizado con fuentes legibles
 */
function generateAndSendPDF(data) {
  try {
    console.log("🚀 Generando PDF PROFESIONAL...");
    
    const htmlContent = generateHTMLProfesionalCorregido(data);
    console.log("✅ HTML PROFESIONAL generado");
    
    // Crear el PDF de forma más robusta
    const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'template.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    
    if (!pdfBlob) {
      throw new Error("No se pudo convertir HTML a PDF");
    }
    
    const fileName = `Entrega_Turno_${data.facturador}_${formatDateForFileName(data.fecha)}_PROFESIONAL.pdf`;
    pdfBlob.setName(fileName);
    
    console.log("📄 PDF PROFESIONAL generado:", fileName);
    
    // Enviar correo con diseño profesional
    const emailResult = sendEmailProfesional(data, pdfBlob);
    
    return {
      success: true,
      emailsSent: emailResult.count,
      fileName: fileName,
      version: PDF_TEMPLATE_VERSION
    };
    
  } catch (error) {
    console.error("❌ Error en generateAndSendPDF:", error);
    return { 
      success: false, 
      error: error.toString(),
      message: "Error al generar PDF"
    };
  }
}

/**
 * Genera HTML PROFESIONAL optimizado para PDF con fuentes legibles
 */
function generateHTMLProfesionalCorregido(data) {
  console.log("🎨 Generando HTML PROFESIONAL CORREGIDO...");
  
  const vehiculosStats = getVehiculosStatsObject(data.vehiculos || []);
  
  return `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        /* ===== RESET Y CONFIGURACIÓN PARA PDF ===== */
        * { 
            margin: 0; 
            padding: 0; 
            box-sizing: border-box; 
            font-family: 'Arial', 'Helvetica', sans-serif;
        }
        
        body {
            font-family: Arial, Helvetica, sans-serif;
            line-height: 1.5;
            color: #333333;
            background: #ffffff;
            padding: 10px;
            font-size: 12pt;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }
        
        .container {
            max-width: 100%;
            margin: 0 auto;
            background: white;
            border: 1px solid #dddddd;
        }
        
        /* ===== HEADER PROFESIONAL ===== */
        .header-profesional {
            background: #2c5aa0;
            color: white;
            padding: 25px 30px;
            text-align: center;
            border-bottom: 5px solid #1e3a8a;
        }
        
        .company-logo {
            font-size: 14pt;
            font-weight: bold;
            margin-bottom: 10px;
            color: #ffffff;
        }
        
        .header-title {
            font-size: 18pt;
            font-weight: bold;
            margin-bottom: 8px;
            color: #ffffff;
        }
        
        .header-subtitle {
            font-size: 11pt;
            opacity: 0.9;
            margin-bottom: 15px;
            color: #e0e7ff;
        }
        
        .header-badges {
            display: flex;
            justify-content: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-top: 15px;
        }
        
        /* ===== ESTADÍSTICAS RÁPIDAS ===== */
        .quick-stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            padding: 20px;
            background: #f8fafc;
            border-bottom: 1px solid #e2e8f0;
        }
        
        .stat-item {
            background: white;
            padding: 15px 10px;
            border-radius: 8px;
            text-align: center;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .stat-icon {
            font-size: 14pt;
            margin-bottom: 5px;
        }
        
        .stat-number {
            font-size: 16pt;
            font-weight: bold;
            color: #1e293b;
            margin-bottom: 3px;
        }
        
        .stat-label {
            font-size: 9pt;
            color: #64748b;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        /* ===== CONTENIDO PRINCIPAL ===== */
        .main-content {
            padding: 20px;
        }
        
        /* ===== SECCIONES ===== */
        .section {
            margin-bottom: 20px;
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 20px;
            page-break-inside: avoid;
        }
        
        .section-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f1f5f9;
        }
        
        .section-icon {
            width: 30px;
            height: 30px;
            background: #2c5aa0;
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12pt;
            color: white;
        }
        
        .section-title {
            font-size: 14pt;
            font-weight: bold;
            color: #1e293b;
            margin: 0;
        }
        
        .section-subtitle {
            font-size: 10pt;
            color: #64748b;
            margin-top: 2px;
        }
        
        /* ===== INFORMACIÓN BÁSICA ===== */
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 12px;
            margin-top: 15px;
        }
        
        .info-card {
            background: #f8fafc;
            padding: 12px;
            border-radius: 6px;
            border-left: 3px solid #2c5aa0;
        }
        
        .info-label {
            font-size: 9pt;
            font-weight: bold;
            color: #2c5aa0;
            text-transform: uppercase;
            margin-bottom: 4px;
            display: block;
        }
        
        .info-value {
            font-size: 11pt;
            font-weight: 600;
            color: #1e293b;
        }
        
        /* ===== TABLAS MEJORADAS ===== */
        .table-container {
            overflow-x: auto;
            margin-top: 15px;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            font-size: 10pt;
        }
        
        .data-table thead {
            background: #2c5aa0;
        }
        
        .data-table th {
            padding: 10px 8px;
            text-align: left;
            font-weight: bold;
            color: white;
            text-transform: uppercase;
            font-size: 9pt;
            border: none;
        }
        
        .data-table td {
            padding: 8px;
            border-bottom: 1px solid #f1f5f9;
            background: white;
        }
        
        .data-table tbody tr:nth-child(even) {
            background: #fafafa;
        }
        
        /* ===== BADGES MEJORADOS ===== */
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 9pt;
            font-weight: bold;
            text-transform: uppercase;
            color: white;
            min-width: 70px;
            text-align: center;
        }
        
        .badge-success { background: #059669; }
        .badge-warning { background: #d97706; }
        .badge-danger { background: #dc2626; }
        .badge-info { background: #0891b2; }
        .badge-primary { background: #2c5aa0; }
        
        /* ===== CONTENIDO DE TEXTO ===== */
        .text-content {
            background: #f0fdf4;
            padding: 15px;
            border-radius: 6px;
            border-left: 3px solid #059669;
            line-height: 1.5;
            color: #065f46;
            font-size: 11pt;
        }
        
        .text-content-warning {
            background: #fffbeb;
            border-left-color: #d97706;
            color: #92400e;
        }
        
        .text-content-danger {
            background: #fef2f2;
            border-left-color: #dc2626;
            color: #991b1b;
        }
        
        /* ===== LISTA DE PUNTOS ===== */
        .puntos-list {
            list-style: none;
            margin-top: 10px;
        }
        
        .punto-item {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            padding: 10px;
            background: #f8fafc;
            margin-bottom: 8px;
            border-radius: 6px;
            border-left: 3px solid #059669;
        }
        
        .punto-icon {
            width: 20px;
            height: 20px;
            background: #059669;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 10pt;
            font-weight: bold;
            flex-shrink: 0;
            margin-top: 1px;
        }
        
        .punto-text {
            flex: 1;
            color: #374151;
            font-weight: 500;
            line-height: 1.4;
            font-size: 11pt;
        }
        
        /* ===== MINI ESTADÍSTICAS ===== */
        .mini-stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 8px;
            margin-bottom: 15px;
        }
        
        .mini-stat {
            background: #f8fafc;
            padding: 10px;
            border-radius: 6px;
            text-align: center;
            border: 1px solid #e2e8f0;
        }
        
        .mini-stat-number {
            font-size: 14pt;
            font-weight: bold;
            color: #1e293b;
            margin-bottom: 2px;
        }
        
        .mini-stat-label {
            font-size: 8pt;
            color: #64748b;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        /* ===== FOOTER ===== */
        .footer {
            background: #1e293b;
            color: white;
            padding: 20px;
            text-align: center;
            margin-top: 20px;
        }
        
        .footer-logo {
            font-size: 12pt;
            font-weight: bold;
            color: #60a5fa;
            margin-bottom: 8px;
        }
        
        .footer-text {
            opacity: 0.8;
            margin-bottom: 5px;
            font-size: 10pt;
        }
        
        .footer-version {
            opacity: 0.6;
            font-size: 9pt;
            font-family: monospace;
        }
        
        /* ===== UTILIDADES ===== */
        .text-center { text-align: center; }
        .mb-2 { margin-bottom: 0.5rem; }
        .mt-2 { margin-top: 0.5rem; }
        .keep-together { page-break-inside: avoid; }
        .page-break { page-break-before: always; }
        
        /* ===== MEJORAS PARA IMPRESIÓN ===== */
        @media print {
            body {
                padding: 5px !important;
                background: white !important;
                font-size: 11pt !important;
            }
            
            .container {
                border: none !important;
                box-shadow: none !important;
                margin: 0 !important;
                max-width: 100% !important;
            }
            
            .section {
                border: 1px solid #cccccc !important;
                box-shadow: none !important;
            }
            
            .header-profesional {
                background: #2c5aa0 !important;
                -webkit-print-color-adjust: exact;
            }
            
            .data-table thead {
                background: #2c5aa0 !important;
                -webkit-print-color-adjust: exact;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- HEADER PROFESIONAL -->
        <div class="header-profesional">
            <div class="company-logo">🚚 COMARRICO LOGÍSTICA</div>
            <h1 class="header-title">ENTREGA DE TURNO - REPORTE OFICIAL</h1>
            <p class="header-subtitle">Sistema de Gestión Logística Empresarial</p>
            <div class="header-badges">
                <span class="badge badge-primary">${formatDateForDisplay(data.fecha)}</span>
                <span class="badge badge-success">${(data.turno || 'NO ESPECIFICADO').toUpperCase()}</span>
                <span class="badge badge-warning">${data.facturador || 'NO ESPECIFICADO'}</span>
            </div>
        </div>
        
        <!-- ESTADÍSTICAS RÁPIDAS -->
        <div class="quick-stats">
            <div class="stat-item">
                <div class="stat-icon">📊</div>
                <div class="stat-number">${vehiculosStats.total}</div>
                <div class="stat-label">Total Vehículos</div>
            </div>
            <div class="stat-item">
                <div class="stat-icon">✅</div>
                <div class="stat-number">${vehiculosStats.confirmados}</div>
                <div class="stat-label">Confirmados</div>
            </div>
            <div class="stat-item">
                <div class="stat-icon">⏳</div>
                <div class="stat-number">${vehiculosStats.pendientes}</div>
                <div class="stat-label">Pendientes</div>
            </div>
            <div class="stat-item">
                <div class="stat-icon">❌</div>
                <div class="stat-number">${vehiculosStats.cancelados}</div>
                <div class="stat-label">Cancelados</div>
            </div>
        </div>
        
        <!-- CONTENIDO PRINCIPAL -->
        <div class="main-content">
            ${buildInformacionBasica(data)}
            ${buildVehiculosCargados(data.vehiculosCargados || [])}
            ${buildObservaciones(data.observaciones)}
            ${buildProyeccionVehiculos(data.vehiculos || [], vehiculosStats)}
            ${buildFacturasPendientes(data.facturasPendientes)}
            ${buildNovedadesMateriales(data.novedadesMateriales)}
            ${buildPuntosCompletados(data.puntosConsiderar || [])}
        </div>
        
        <!-- FOOTER -->
        <div class="footer">
            <div class="footer-logo">COMARRICO LOGÍSTICA</div>
            <p class="footer-text">Sistema de Gestión Logística Empresarial</p>
            <p class="footer-text">Reporte generado: ${new Date().toLocaleString('es-ES')}</p>
            <p class="footer-version">v${PDF_TEMPLATE_VERSION}</p>
        </div>
    </div>
</body>
</html>
  `;
}

/**
 * COMPONENTES DEL PDF PROFESIONAL
 */

function buildInformacionBasica(data) {
  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">👤</div>
            <div>
                <h2 class="section-title">Información del Turno</h2>
                <p class="section-subtitle">Detalles principales de la entrega</p>
            </div>
        </div>
        <div class="info-grid">
            <div class="info-card">
                <span class="info-label">Facturador</span>
                <div class="info-value">${data.facturador || 'No especificado'}</div>
            </div>
            <div class="info-card">
                <span class="info-label">Turno</span>
                <div class="info-value">${(data.turno || 'No especificado').toUpperCase()}</div>
            </div>
            <div class="info-card">
                <span class="info-label">Fecha</span>
                <div class="info-value">${formatDateForDisplay(data.fecha)}</div>
            </div>
            <div class="info-card">
                <span class="info-label">Hora de Reporte</span>
                <div class="info-value">${new Date().toLocaleTimeString('es-ES', { 
                    hour: '2-digit', 
                    minute: '2-digit',
                    second: '2-digit'
                })}</div>
            </div>
        </div>
    </div>
  `;
}

function buildVehiculosCargados(vehiculosCargados) {
  if (!vehiculosCargados || vehiculosCargados.length === 0) {
    return '';
  }

  const rows = vehiculosCargados.map((vehiculo, index) => {
    return `
      <tr>
        <td><strong>${index + 1}</strong></td>
        <td>${vehiculo.destino || '-'}</td>
        <td>${vehiculo.transportadora || '-'}</td>
        <td><strong>${vehiculo.placa || '-'}</strong></td>
        <td>${vehiculo.quienCargo || '-'}</td>
        <td>${vehiculo.quienPreparo || '-'}</td>
        <td><span class="badge badge-success">${vehiculo.estatus || 'COMPLETADO'}</span></td>
      </tr>
    `;
  }).join('');

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">🚛</div>
            <div>
                <h2 class="section-title">Vehículos Cargados</h2>
                <p class="section-subtitle">Durante el turno actual</p>
            </div>
        </div>
        <div class="table-container">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Destino</th>
                        <th>Transportadora</th>
                        <th>Placa</th>
                        <th>Verificador</th>
                        <th>Separador</th>
                        <th>Estatus</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
    </div>
  `;
}

function buildObservaciones(observaciones) {
  if (!observaciones || observaciones.trim() === '') {
    return '';
  }

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">📝</div>
            <div>
                <h2 class="section-title">Observaciones</h2>
                <p class="section-subtitle">Puntos importantes del turno</p>
            </div>
        </div>
        <div class="text-content">
            ${observaciones.replace(/\n/g, '<br>')}
        </div>
    </div>
  `;
}

function buildProyeccionVehiculos(vehiculos, stats) {
  const rows = vehiculos.map((vehiculo, index) => {
    const estadoClass = getEstadoClass(vehiculo.estado);
    const estadoText = getEstadoText(vehiculo.estado);
    
    return `
      <tr>
        <td><strong>${index + 1}</strong></td>
        <td>${vehiculo.destino || '-'}</td>
        <td>${vehiculo.transportadora || '-'}</td>
        <td><strong>${vehiculo.placa || '-'}</strong></td>
        <td>${vehiculo.horaCitada || '-'}</td>
        <td><span class="badge ${estadoClass}">${estadoText}</span></td>
        <td>${vehiculo.consecutivo || '-'}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">📊</div>
            <div>
                <h2 class="section-title">Proyección de Vehículos</h2>
                <p class="section-subtitle">Programación del día siguiente</p>
            </div>
        </div>
        
        <div class="mini-stats">
            <div class="mini-stat">
                <div class="mini-stat-number">${stats.total}</div>
                <div class="mini-stat-label">Total</div>
            </div>
            <div class="mini-stat">
                <div class="mini-stat-number">${stats.confirmados}</div>
                <div class="mini-stat-label">Confirmados</div>
            </div>
            <div class="mini-stat">
                <div class="mini-stat-number">${stats.pendientes}</div>
                <div class="mini-stat-label">Pendientes</div>
            </div>
            <div class="mini-stat">
                <div class="mini-stat-number">${stats.cancelados}</div>
                <div class="mini-stat-label">Cancelados</div>
            </div>
        </div>
        
        <div class="table-container">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Destino</th>
                        <th>Transportadora</th>
                        <th>Placa</th>
                        <th>Hora</th>
                        <th>Estado</th>
                        <th>Cons.</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows || '<tr><td colspan="7" class="text-center">No hay vehículos programados</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
  `;
}

function buildFacturasPendientes(facturasPendientes) {
  if (!facturasPendientes || facturasPendientes.trim() === '') {
    return '';
  }

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">💰</div>
            <div>
                <h2 class="section-title">Facturas Pendientes</h2>
                <p class="section-subtitle">Fletes y novedades financieras</p>
            </div>
        </div>
        <div class="text-content text-content-warning">
            ${facturasPendientes.replace(/\n/g, '<br>')}
        </div>
    </div>
  `;
}

function buildNovedadesMateriales(novedadesMateriales) {
  if (!novedadesMateriales || novedadesMateriales.trim() === '') {
    return '';
  }

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">⚠️</div>
            <div>
                <h2 class="section-title">Novedades Materiales</h2>
                <p class="section-subtitle">Incidencias y observaciones</p>
            </div>
        </div>
        <div class="text-content text-content-danger">
            ${novedadesMateriales.replace(/\n/g, '<br>')}
        </div>
    </div>
  `;
}

function buildPuntosCompletados(puntosConsiderar) {
  if (!puntosConsiderar || puntosConsiderar.length === 0) {
    return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">✅</div>
            <div>
                <h2 class="section-title">Actividades Completadas</h2>
                <p class="section-subtitle">Puntos considerados en el turno</p>
            </div>
        </div>
        <div class="text-center" style="padding: 20px; color: #64748b;">
            <div style="font-size: 16pt; margin-bottom: 10px;">📋</div>
            <div style="font-weight: 600;">No se registraron actividades específicas</div>
        </div>
    </div>
    `;
  }

  const puntosContent = puntosConsiderar.map((punto, index) => {
    return `
      <li class="punto-item">
        <div class="punto-icon">${index + 1}</div>
        <div class="punto-text">${punto}</div>
      </li>
    `;
  }).join('');

  return `
    <div class="section keep-together">
        <div class="section-header">
            <div class="section-icon">✅</div>
            <div>
                <h2 class="section-title">Actividades Completadas</h2>
                <p class="section-subtitle">${puntosConsiderar.length} puntos considerados</p>
            </div>
        </div>
        <ul class="puntos-list">
            ${puntosContent}
        </ul>
    </div>
  `;
}

/**
 * FUNCIONES AUXILIARES MEJORADAS
 */

function getVehiculosStatsObject(vehiculos) {
  if (!vehiculos || !Array.isArray(vehiculos)) {
    return { total: 0, confirmados: 0, pendientes: 0, cancelados: 0 };
  }

  const stats = {
    total: vehiculos.length,
    confirmados: 0,
    pendientes: 0,
    cancelados: 0
  };

  vehiculos.forEach(vehiculo => {
    const estado = (vehiculo.estado || 'pendiente').toLowerCase();
    if (estado === 'confirmado') {
      stats.confirmados++;
    } else if (estado === 'cancelado') {
      stats.cancelados++;
    } else {
      stats.pendientes++;
    }
  });

  return stats;
}

function formatDateForDisplay(dateString) {
  if (!dateString) return 'Fecha no especificada';
  try {
    const date = new Date(dateString);
    return date.toLocaleDateString('es-ES', {
      weekday: 'long',
      year: 'numeric', 
      month: 'long',
      day: 'numeric'
    });
  } catch (error) {
    return dateString;
  }
}

function formatDateForFileName(dateString) {
  if (!dateString) return 'fecha-no-especificada';
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  } catch (error) {
    return dateString.replace(/-/g, '');
  }
}

function getEstadoClass(estado) {
  const estadoMap = {
    'confirmado': 'badge-success',
    'pendiente': 'badge-warning',
    'cancelado': 'badge-danger'
  };
  return estadoMap[estado] || 'badge-warning';
}

function getEstadoText(estado) {
  const textMap = {
    'confirmado': 'CONFIRMADO',
    'pendiente': 'PENDIENTE',
    'cancelado': 'CANCELADO'
  };
  return textMap[estado] || 'PENDIENTE';
}

/**
 * ENVÍO DE CORREO PROFESIONAL
 */
function sendEmailProfesional(data, pdfBlob) {
  try {
    const fechaFormateada = formatDateForDisplay(data.fecha);
    const subject = `📋 Entrega de Turno - ${data.facturador} - ${fechaFormateada.split(',')[0]}`;
    
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f8fafc;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
        .header {
            background: linear-gradient(135deg, #2c5aa0, #1e3a8a);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 24px;
            font-weight: bold;
        }
        .content {
            padding: 30px;
        }
        .info-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin: 20px 0;
        }
        .info-item {
            background: #f8fafc;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #2c5aa0;
        }
        .info-label {
            font-size: 12px;
            color: #64748b;
            font-weight: bold;
            text-transform: uppercase;
        }
        .info-value {
            font-size: 14px;
            font-weight: 600;
            color: #1e293b;
            margin-top: 5px;
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            margin: 25px 0;
        }
        .stat {
            text-align: center;
            padding: 15px;
            background: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
        }
        .stat-number {
            font-size: 20px;
            font-weight: bold;
            color: #2c5aa0;
        }
        .stat-label {
            font-size: 11px;
            color: #64748b;
            text-transform: uppercase;
            font-weight: 600;
            margin-top: 5px;
        }
        .attachment {
            background: #dcfce7;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #059669;
            margin: 20px 0;
        }
        .footer {
            background: #1e293b;
            color: white;
            padding: 25px;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚚 ENTREGA DE TURNO REGISTRADA</h1>
            <p style="margin: 10px 0 0 0; opacity: 0.9;">Sistema de Gestión Logística</p>
        </div>
        
        <div class="content">
            <h2 style="color: #1e293b; margin-bottom: 10px;">Nuevo reporte de entrega de turno</h2>
            <p style="color: #64748b; margin-bottom: 25px;">
                Se ha registrado exitosamente una nueva entrega de turno en el sistema.
            </p>
            
            <div class="info-grid">
                <div class="info-item">
                    <div class="info-label">Facturador</div>
                    <div class="info-value">${data.facturador}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Turno</div>
                    <div class="info-value">${data.turno.toUpperCase()}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Fecha</div>
                    <div class="info-value">${fechaFormateada}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Hora</div>
                    <div class="info-value">${new Date().toLocaleTimeString('es-ES')}</div>
                </div>
            </div>
            
            <div class="stats">
                <div class="stat">
                    <div class="stat-number">${data.vehiculos ? data.vehiculos.length : 0}</div>
                    <div class="stat-label">Vehículos</div>
                </div>
                <div class="stat">
                    <div class="stat-number">${data.puntosConsiderar ? data.puntosConsiderar.length : 0}</div>
                    <div class="stat-label">Actividades</div>
                </div>
                <div class="stat">
                    <div class="stat-number">${data.observaciones ? 'SÍ' : 'NO'}</div>
                    <div class="stat-label">Observaciones</div>
                </div>
                <div class="stat">
                    <div class="stat-number">${data.facturasPendientes ? 'SÍ' : 'NO'}</div>
                    <div class="stat-label">Facturas Pend.</div>
                </div>
            </div>
            
            <div class="attachment">
                <h3 style="color: #065f46; margin: 0 0 10px 0;">📎 Reporte Adjunto</h3>
                <p style="margin: 0; color: #065f46;">
                    <strong>Archivo:</strong> ${pdfBlob.getName()}<br>
                    <strong>Tamaño:</strong> ${Math.round(pdfBlob.getBytes().length / 1024)} KB
                </p>
            </div>
            
            <p style="color: #475569; font-size: 14px;">
                Este reporte contiene información detallada sobre la entrega de turno, incluyendo 
                proyección de vehículos, observaciones y actividades completadas.
            </p>
        </div>
        
        <div class="footer">
            <p style="margin: 0; opacity: 0.8;">Sistema de Gestión Logística Empresarial</p>
            <p style="margin: 5px 0 0 0; opacity: 0.6; font-size: 12px;">
                Generado automáticamente • ${new Date().toLocaleString('es-ES')}
            </p>
        </div>
    </div>
</body>
</html>
    `;
    
    const textBody = `
ENTREGA DE TURNO REGISTRADA - COMARRICO LOGÍSTICA

Nuevo reporte de entrega de turno registrado en el sistema:

📊 INFORMACIÓN PRINCIPAL:
• Facturador: ${data.facturador}
• Turno: ${data.turno.toUpperCase()}
• Fecha: ${fechaFormateada}
• Hora: ${new Date().toLocaleTimeString('es-ES')}

📈 ESTADÍSTICAS:
• Vehículos programados: ${data.vehiculos ? data.vehiculos.length : 0}
• Actividades completadas: ${data.puntosConsiderar ? data.puntosConsiderar.length : 0}
• Observaciones: ${data.observaciones ? 'SÍ' : 'NO'}
• Facturas pendientes: ${data.facturasPendientes ? 'SÍ' : 'NO'}

📎 ARCHIVO ADJUNTO:
Se adjunta el reporte completo en formato PDF: ${pdfBlob.getName()}

---
Este es un mensaje automático del Sistema de Gestión Logística.
Generado el ${new Date().toLocaleString('es-ES')}
    `.trim();

    let emailsSent = 0;
    
    // Enviar a correos de notificación
    CORREOS_NOTIFICACION.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: textBody,
          htmlBody: htmlBody,
          attachments: [pdfBlob],
          name: 'Sistema de Entrega de Turno - Comarrico'
        });
        emailsSent++;
        console.log(`✅ Correo profesional enviado a: ${email}`);
      } catch (emailError) {
        console.error(`❌ Error enviando a ${email}:`, emailError);
      }
    });

    return {
      success: true,
      count: emailsSent,
      recipients: CORREOS_NOTIFICACION
    };
    
  } catch (error) {
    console.error("❌ Error en sendEmailProfesional:", error);
    return {
      success: false,
      error: error.toString(),
      count: 0
    };
  }
}

/**
 * Función de prueba PROFESIONAL
 */
function testPDFProfesional() {
  console.log("🧪 Probando PDF PROFESIONAL...");
  
  const testData = {
    facturador: "JUAN CARLOS GÓMEZ",
    turno: "mañana",
    fecha: new Date().toISOString().split('T')[0],
    observaciones: "Reporte de prueba con el NUEVO DISEÑO PROFESIONAL. Fuentes legibles y diseño optimizado para PDF.",
    vehiculos: [
      {
        destino: "BOGOTÁ D.C. - CENTRO DISTRIBUCIÓN NORTE",
        transportadora: "TRANSPORTES ÉLITE S.A.S.",
        placa: "ABC123",
        horaCitada: "08:00",
        estado: "confirmado",
        consecutivo: "CT-001"
      },
      {
        destino: "MEDELLÍN - ZONA INDUSTRIAL SUR",
        transportadora: "CARGA EXPRESS COLOMBIA",
        placa: "XYZ789",
        horaCitada: "09:30",
        estado: "pendiente",
        consecutivo: "CT-002"
      }
    ],
    vehiculosCargados: [
      {
        destino: "MEDELLÍN - ZONA INDUSTRIAL",
        transportadora: "CARGA RÁPIDA EXPRESS LTDA",
        placa: "XYZ789",
        quienCargo: "CARLOS RAMÍREZ",
        quienPreparo: "MARÍA LÓPEZ", 
        estatus: "terminado"
      }
    ],
    facturasPendientes: "Factura #FV-2024-001 pendiente de aprobación. Factura #FV-2024-003 en verificación.",
    novedadesMateriales: "Retraso en material de empaque. Coordinar con proveedor.",
    puntosConsiderar: [
      "Revisión completa de facturas de fletes mensuales",
      "Solicitud de documentación para despacho internacional",
      "Coordinación con transportadoras para programación",
      "Verificación de inventario de materiales",
      "Actualización de base de datos"
    ]
  };
  
  const result = generateAndSendPDF(testData);
  console.log("🎯 Resultado PROFESIONAL:", result);
  return result;
}