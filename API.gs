/**
 * ============================================================
 * SIBRU - API Endpoints
 * Sistema de Tarjetas Digitales
 * Backend para GitHub Pages + Apps Script
 * ============================================================
 */

// ============================================================
// CONFIGURACIÓN GLOBAL
// ============================================================
const CONFIG = {
  SPREADSHEET_ID: '1SkcZljYKuHYvGf9nrdr6V28aBhmG5COkfhhzEg6xlHQ',
  SHEETS: {
    CONFIG: 'Config',
    EMPLEADOS: 'Empleados',
    GALERIA: 'Galeria',
    LINKS: 'Links',
    LEADS: 'Leads'
  },
  APP: {
    DEFAULT_CARD_ID: 'empresa',
    DEFAULT_LOGO: 'https://res.cloudinary.com/sdhnyr5e/image/upload/v1793814011/Retrato_Google_con_Sonrisa.png'
  }
};

// ============================================================
// PUNTO DE ENTRADA - GET
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const cardId = params.id || CONFIG.APP.DEFAULT_CARD_ID;
    
    // Si solicita JSON (API para GitHub Pages)
    if (params.format === 'json') {
      return handleAPIRequest(params);
    }
    
    // Si solicita panel admin
    if (params.admin && String(params.admin) === 'true') {
      const template = HtmlService.createTemplateFromFile('admin');
      return template.evaluate()
        .setTitle('Panel Administrativo SIBRU')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // Si solicita tarjeta personal
    if (params.tipo && String(params.tipo) === 'personal') {
      return generarHTMLPersonal(cardId);
    }
    
    // Por defecto: tarjeta corporativa
    const template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('Grupo Sibrü | Tarjeta Digital')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// PUNTO DE ENTRADA - POST
// ============================================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    return handlePostRequest(params);
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// MANEJO DE PETICIONES GET (API JSON)
// ============================================================
function handleAPIRequest(params) {
  const action = params.action || 'getCard';
  const id = params.id || CONFIG.APP.DEFAULT_CARD_ID;
  
  try {
    let data;
    
    switch(action) {
      case 'getCard':
        data = getCardData(id);
        break;
      case 'getConfig':
        data = { success: true, data: getConfig() };
        break;
      case 'getAllEmployees':
        data = { success: true, data: getEmpleados(false) };
        break;
      case 'getEmployee':
        data = { success: true, data: getEmpleadoById(id) };
        break;
      case 'getGallery':
        data = { success: true, data: getGaleriaByEmpleado(id) };
        break;
      case 'getLinks':
        data = { success: true, data: getLinksByEmpleado(id) };
        break;
      case 'validatePin':
        data = validateAdminPin(params.pin);
        break;
      default:
        data = { success: false, error: 'Acción no válida: ' + action };
    }
    
    return jsonResponse(data);
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// MANEJO DE PETICIONES POST
// ============================================================
function handlePostRequest(params) {
  const action = params.action;
  
  try {
    let data;
    
    switch(action) {
      case 'validatePin':
        data = validateAdminPin(params.pin);
        break;
      case 'addEmployee':
        data = addEmpleado(params.employee);
        break;
      case 'updateEmployee':
        data = updateEmpleado(params.id, params.updates);
        break;
      case 'deleteEmployee':
        data = deleteEmpleado(params.id);
        break;
      case 'updateGallery':
        data = updateGaleriaEmpleado(params.id, params.images);
        break;
      case 'updateLinks':
        data = updateLinksEmpleado(params.id, params.links);
        break;
      case 'saveLead':
        data = guardarLead(params.lead || params);
        break;
      default:
        data = { success: false, error: 'Acción no válida: ' + action };
    }
    
    return jsonResponse(data);
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// FUNCIONES DE BASE DE DATOS - EMPLEADOS
// ============================================================

function getEmpleados(activos) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLEADOS);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const empleados = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    return activos !== false ? empleados.filter(emp => emp.ACTIVO === true || emp.ACTIVO === 'TRUE') : empleados;
  } catch (error) {
    Logger.log('Error getEmpleados: ' + error.message);
    return [];
  }
}

function getEmpleadoById(id) {
  try {
    const empleados = getEmpleados(false);
    const empleado = empleados.find(emp => emp.ID === id);
    return empleado || null;
  } catch (error) {
    Logger.log('Error getEmpleadoById: ' + error.message);
    return null;
  }
}

function addEmpleado(employee) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLEADOS);
    
    if (!sheet) {
      return { success: false, error: 'Hoja Empleados no encontrada' };
    }
    
    // Verificar si ya existe el ID
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === employee.ID) {
        return { success: false, error: 'Ya existe un empleado con ese ID' };
      }
    }
    
    // Agregar nuevo empleado
    const newRow = [
      employee.ID || '',
      employee.NOMBRE || '',
      employee.CARGO || '',
      employee.TELEFONO || '',
      employee.EMAIL || '',
      employee.FOTO_PERFIL || '',
      employee.ACTIVO !== undefined ? employee.ACTIVO : true,
      employee.TIPO || 'empleado'
    ];
    
    sheet.appendRow(newRow);
    
    return { 
      success: true, 
      message: 'Empleado agregado correctamente', 
      id: employee.ID 
    };
  } catch (error) {
    Logger.log('Error addEmpleado: ' + error.message);
    return { success: false, error: error.message };
  }
}

function updateEmpleado(id, updates) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLEADOS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const rowIndex = data.findIndex(row => row[headers.indexOf('ID')] === id);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Empleado no encontrado' };
    }
    
    Object.keys(updates).forEach(key => {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(updates[key]);
      }
    });
    
    return { success: true, message: 'Empleado actualizado' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteEmpleado(id) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLEADOS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const rowIndex = data.findIndex(row => row[headers.indexOf('ID')] === id);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Empleado no encontrado' };
    }
    
    sheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'Empleado eliminado' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================================
// FUNCIONES DE BASE DE DATOS - CONFIGURACIÓN
// ============================================================

function getConfig() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return {};
    }
    
    const data = sheet.getDataRange().getValues();
    const config = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        config[data[i][0]] = data[i][1];
      }
    }
    
    return config;
  } catch (error) {
    Logger.log('Error getConfig: ' + error.message);
    return {};
  }
}

// ============================================================
// FUNCIONES DE BASE DE DATOS - GALERÍA Y LINKS
// ============================================================

function getGaleriaByEmpleado(idEmpleado) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.GALERIA);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const galeria = data.slice(1)
      .filter(row => row[headers.indexOf('ID_EMPLEADO')] === idEmpleado && (row[headers.indexOf('ACTIVO')] === true || row[headers.indexOf('ACTIVO')] === 'TRUE'))
      .map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      })
      .sort((a, b) => a.ORDEN - b.ORDEN);
    
    return galeria;
  } catch (error) {
    return [];
  }
}

function getLinksByEmpleado(idEmpleado) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LINKS);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const links = data.slice(1)
      .filter(row => row[headers.indexOf('ID_EMPLEADO')] === idEmpleado && (row[headers.indexOf('ACTIVO')] === true || row[headers.indexOf('ACTIVO')] === 'TRUE'))
      .map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      })
      .sort((a, b) => a.ORDEN - b.ORDEN);
    
    return links;
  } catch (error) {
    return [];
  }
}

function getCardData(idEmpleado) {
  try {
    const empleado = getEmpleadoById(idEmpleado);
    
    if (!empleado) {
      return { success: false, error: 'Empleado no encontrado' };
    }
    
    const config = getConfig();
    const galeria = getGaleriaByEmpleado(idEmpleado);
    const links = getLinksByEmpleado(idEmpleado);
    
    return {
      success: true,
      data: {
        config: config,
        empleado: empleado,
        galeria: galeria,
        links: links
      }
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================================
// FUNCIONES DE BASE DE DATOS - LEADS
// ============================================================

function guardarLead(leadData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEETS.LEADS);
    
    // Crear hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.LEADS);
      sheet.appendRow(['ID', 'FECHA', 'NOMBRE', 'TELEFONO', 'EMAIL', 'TIPO_SOLICITUD', 'PRODUCTOS_INTERES', 'MENSAJE', 'ESTADO']);
    }
    
    const leadId = 'LEAD' + Utilities.formatDate(new Date(), 'GMT-6', 'yyyyMMddHHmmss');
    
    sheet.appendRow([
      leadId,
      new Date(),
      leadData.nombre || '',
      leadData.telefono || '',
      leadData.email || '',
      leadData.tipoSolicitud || '',
      leadData.productosInteres || '',
      leadData.mensaje || '',
      'Pendiente'
    ]);
    
    return { 
      success: true, 
      message: 'Lead guardado exitosamente',
      leadId: leadId
    };
  } catch (error) {
    Logger.log('Error guardarLead: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ============================================================
// FUNCIONES DE SEGURIDAD
// ============================================================

function validateAdminPin(pin) {
  try {
    const config = getConfig();
    const isValid = String(pin) === String(config.admin_pin);
    return { success: isValid };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================================
// FUNCIONES PUENTE (para google.script.run desde HTML en Apps Script)
// ============================================================

function getCardDataForClient(id) {
  try {
    const cardId = id || CONFIG.APP.DEFAULT_CARD_ID;
    return getCardData(cardId);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function submitLeadForm(leadData) {
  return guardarLead(leadData);
}

// ============================================================
// GENERACIÓN DE HTML PERSONAL (fallback si no se usa GitHub Pages)
// ============================================================

function generarHTMLPersonal(cardId) {
  try {
    const empleado = getEmpleadoById(cardId);
    const config = getConfig();
    
    if (!empleado) {
      return HtmlService.createHtmlOutput('<h1 style="text-align:center;padding:40px;">Empleado no encontrado</h1>')
        .setTitle('Error')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    const name = empleado.NOMBRE || 'Sin nombre';
    const cargo = empleado.CARGO || config.empresa_tagline || '';
    const phone = empleado.TELEFONO || config.empresa_whatsapp || '';
    const phoneClean = String(phone).replace(/^(\+?52)?/, '').trim();
    const phoneFmt = phoneClean.length === 10 ? phoneClean.slice(0,2) + ' ' + phoneClean.slice(2,6) + ' ' + phoneClean.slice(6) : phoneClean;
    const email = empleado.EMAIL || config.empresa_email || '';
    const fotoPerfil = empleado.FOTO_PERFIL || '';
    const nuevoLogo = config.empresa_logo || CONFIG.APP.DEFAULT_LOGO;
    const baseUrl = ScriptApp.getService().getUrl();
    
    const initials = name.split(' ').map(function(n) { return n.charAt(0); }).join('').substring(0, 2).toUpperCase();
    
    let avatarHTML = '<span style="font-size:40px;font-weight:700;color:#8BC34A;">' + initials + '</span>';
    if (fotoPerfil && fotoPerfil.trim() !== '') {
      avatarHTML = '<img src="' + fotoPerfil + '" alt="' + name + '" style="width:100%;height:100%;object-fit:cover;">';
    }
    
    let html = '<!DOCTYPE html>';
    html += '<html lang="es" prefix="og: http://ogp.me/ns#">';
    html += '<head>';
    html += '<meta charset="UTF-8">';
    html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
    html += '<meta name="theme-color" content="#000000">';
    html += '<meta property="og:type" content="profile">';
    html += '<meta property="og:url" content="' + baseUrl + '?tipo=personal&id=' + cardId + '">';
    html += '<meta property="og:title" content="' + name + ' - ' + (config.empresa_nombre || '') + '">';
    html += '<meta property="og:description" content="' + cargo + ' | ' + (config.empresa_ubicacion || '') + '">';
    html += '<meta property="og:image" content="' + (fotoPerfil || nuevoLogo) + '">';
    html += '<meta property="og:image:width" content="1200">';
    html += '<meta property="og:image:height" content="630">';
    html += '<meta property="og:site_name" content="' + (config.empresa_nombre || 'Grupo Sibrü') + '">';
    html += '<meta name="twitter:card" content="summary_large_image">';
    html += '<meta name="twitter:title" content="' + name + '">';
    html += '<meta name="twitter:description" content="' + cargo + '">';
    html += '<meta name="twitter:image" content="' + (fotoPerfil || nuevoLogo) + '">';
    html += '<title>' + name + ' | ' + (config.empresa_nombre || '') + '</title>';
    html += '<style>';
    html += '* { margin: 0; padding: 0; box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }';
    html += 'body { background: #000000; min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px; }';
    html += '.card { background: #ffffff; border-radius: 24px; padding: 40px 28px; max-width: 380px; width: 100%; box-shadow: 0 20px 60px rgba(139, 195, 74, 0.2); text-align: center; }';
    html += '.logo-container { margin-bottom: 24px; }';
    html += '.logo-large { width: 100px; height: 100px; margin: 0 auto; display: block; border-radius: 16px; box-shadow: 0 8px 24px rgba(139, 195, 74, 0.3); }';
    html += '.avatar { width: 110px; height: 110px; border-radius: 50%; background: linear-gradient(135deg, #8BC34A, #33691E); margin: 0 auto 20px; display: flex; align-items: center; justify-content: center; border: 4px solid #8BC34A; overflow: hidden; box-shadow: 0 8px 24px rgba(139, 195, 74, 0.4); }';
    html += '.name { font-size: 26px; font-weight: 700; color: #1a1a1a; margin-bottom: 6px; }';
    html += '.role { font-size: 15px; color: #666; margin-bottom: 8px; font-weight: 500; }';
    html += '.company { font-size: 13px; color: #8BC34A; margin-bottom: 24px; font-weight: 600; }';
    html += '.actions { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 16px; }';
    html += '.action-btn { background: #f8f9fa; border: 2px solid #8BC34A; border-radius: 14px; padding: 14px 10px; text-decoration: none; color: #1a1a1a; display: flex; flex-direction: column; align-items: center; gap: 8px; transition: all 0.3s; cursor: pointer; font-family: inherit; font-size:12px; font-weight: 600; }';
    html += '.action-btn:hover { background: #8BC34A; transform: translateY(-3px); box-shadow: 0 6px 16px rgba(139, 195, 74, 0.3); }';
    html += '.action-btn:hover svg { fill: #fff; }';
    html += '.action-btn svg { width: 26px; height: 26px; fill: #8BC34A; transition: all 0.3s; }';
    html += '.corporate-btn { grid-column: span 2; background: linear-gradient(135deg, #8BC34A, #689F38); border: none; border-radius: 14px; padding: 16px; color: #fff; text-decoration: none; display: flex; flex-direction: column; align-items: center; gap: 8px; transition: all 0.3s; cursor: pointer; font-family: inherit; font-size:14px; font-weight: 700; margin-top: 8px; }';
    html += '.corporate-btn:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(139, 195, 74, 0.4); }';
    html += '.corporate-btn svg { width: 28px; height: 28px; fill: #fff; }';
    html += '.divider { height: 1px; background: linear-gradient(90deg, transparent, #8BC34A, transparent); margin: 20px 0; }';
    html += '.contact-info { font-size: 13px; color: #666; line-height: 2; }';
    html += '.contact-info a { color: #8BC34A; text-decoration: none; font-weight: 600; }';
    html += '.footer-brand { margin-top: 20px; font-size: 11px; color: #999; }';
    html += '</style>';
    html += '</head><body>';
    html += '<div class="card">';
    html += '<div class="logo-container"><img src="' + nuevoLogo + '" alt="' + (config.empresa_nombre || 'Logo') + '" class="logo-large"></div>';
    html += '<div class="avatar">' + avatarHTML + '</div>';
    html += '<div class="name">' + name + '</div>';
    html += '<div class="role">' + cargo + '</div>';
    html += '<div class="company">' + (config.empresa_nombre || 'Grupo Sibrü') + '</div>';
    html += '<div class="actions">';
    html += '<a href="https://wa.me/52' + phoneClean + '?text=' + encodeURIComponent(config.empresa_whatsapp_msg || 'Hola') + '" target="_blank" class="action-btn"><svg viewBox="0 0 24 24"><path d="M20 2H4c-1.1 0-2 .9-2 2v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2z"/></svg><span>WhatsApp</span></a>';
    html += '<a href="tel:' + phoneClean + '" class="action-btn"><svg viewBox="0 0 24 24"><path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/></svg><span>Llamar</span></a>';
    html += '<a href="mailto:' + email + '" class="action-btn"><svg viewBox="0 0 24 24"><path d="M20 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg><span>Email</span></a>';
    html += '<a href="#" onclick="downloadVCF(); return false;" class="action-btn"><svg viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg><span>Guardar</span></a>';
    html += '<a href="' + baseUrl + '" target="_blank" class="corporate-btn"><svg viewBox="0 0 24 24"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/></svg><span>Ver Información de la Empresa</span></a>';
    html += '</div>';
    html += '<div class="divider"></div>';
    html += '<div class="contact-info">';
    html += '<div>📞 ' + phoneFmt + '</div>';
    html += '<div>✉️ <a href="mailto:' + email + '">' + email + '</a></div>';
    html += '</div>';
    html += '<div class="footer-brand">' + (config.empresa_nombre || 'Grupo Sibrü') + ' © 2026</div>';
    html += '</div>';
    html += '<script>';
    html += 'function downloadVCF() { var vcf = "BEGIN:VCARD\\r\\nVERSION:3.0\\r\\nFN:' + name + '\\r\\nORG:' + (config.empresa_nombre || '') + '\\r\\nTITLE:' + cargo + '\\r\\nTEL;TYPE=CELL:+52' + phone + '\\r\\nEMAIL:' + email + '\\r\\nEND:VCARD"; var blob = new Blob([vcf], { type: "text/vcard;charset=utf-8" }); var url = URL.createObjectURL(blob); var a = document.createElement("a"); a.href = url; a.download = "' + name.replace(/\s+/g, '_') + '.vcf"; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url); }';
    html += '</script>';
    html += '</body></html>';
    
    return HtmlService.createHtmlOutput(html)
      .setMimeType(HtmlService.MimeType.HTML)
      .setTitle(name)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    Logger.log('Error generarHTMLPersonal: ' + error.message);
    return HtmlService.createHtmlOutput('<h1>Error al generar tarjeta</h1>')
      .setMimeType(HtmlService.MimeType.HTML);
  }
}

// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleError(error, action) {
  Logger.log('Error en ' + action + ': ' + error.message);
  return {
    success: false,
    error: error.message,
    action: action
  };
}

// ============================================================
// FUNCIONES DE ACTUALIZACIÓN DE GALERÍA Y LINKS
// ============================================================

function updateGaleriaEmpleado(idEmpleado, imagenes) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.GALERIA);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Eliminar imágenes existentes del empleado
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][headers.indexOf('ID_EMPLEADO')] === idEmpleado) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // Agregar nuevas imágenes
    imagenes.forEach((img, index) => {
      sheet.appendRow([
        idEmpleado,
        img.url || img.IMAGEN_URL || '',
        img.descripcion || img.DESCRIPCION || '',
        index + 1,
        true
      ]);
    });
    
    return { success: true, message: 'Galería actualizada' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateLinksEmpleado(idEmpleado, links) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LINKS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Eliminar links existentes del empleado
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][headers.indexOf('ID_EMPLEADO')] === idEmpleado) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // Agregar nuevos links
    links.forEach((link, index) => {
      sheet.appendRow([
        idEmpleado,
        link.titulo || link.TITULO || '',
        link.subtitulo || link.SUBTITULO || '',
        link.url || link.URL || '',
        link.icono || link.ICONO || 'link',
        index + 1,
        true
      ]);
    });
    
    return { success: true, message: 'Links actualizados' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
