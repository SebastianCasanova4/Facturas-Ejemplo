function onOpen() {
  showSidebar();  // Mostrar la barra lateral automáticamente al abrir el documento
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Show sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('main')
      .setTitle('Menú prueba');
      SpreadsheetApp.getUi()
      .showSidebar(html);
    }
function showPreProductos() {
  var html = HtmlService.createHtmlOutputFromFile('preProductos')
    .setTitle('Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showAggProductos() {
  var html = HtmlService.createHtmlOutputFromFile('agregarProducto')
    .setTitle('Agregar Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}
    
function openFacturaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PreFactura");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showMenuFactura() {
  var html = HtmlService.createHtmlOutputFromFile('menuFactura')
    .setTitle('Menú Factura');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openClientesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  SpreadsheetApp.setActiveSheet(sheet);
}

function openProductosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Productos");
  SpreadsheetApp.setActiveSheet(sheet);ç
}

function showEnviarEmail() {
  var html = HtmlService.createHtmlOutputFromFile('enviarEmail')
    .setTitle('Enviar Email');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function processForm(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const codigoReferencia = data.codigoReferencia;
    const nombre = data.nombre;
    const valorUnitario = parseFloat(data.valorUnitario);
    const iva = parseFloat(data.iva) / 100;
    const precioConIva = valorUnitario * (1 + iva);
    const impuestos = valorUnitario * iva;
    
    sheet.getRange(newRow, 1).setValue(codigoReferencia);
    sheet.getRange(newRow, 2).setValue(nombre);
    sheet.getRange(newRow, 3).setValue(valorUnitario);
    // Establece el IVA y formatea la celda como porcentaje
    const ivaCell = sheet.getRange(newRow, 4);
    ivaCell.setValue(iva); // Establece el valor del IVA como decimal
    ivaCell.setNumberFormat('0.00%'); // Formatea la celda como porcentaje con dos decimales
    sheet.getRange(newRow, 5).setValue(precioConIva); // Guarda el precio con IVA
    sheet.getRange(newRow, 6).setValue(impuestos); // Guarda el valor de los impuestos
    
    

    return "Datos guardados correctamente";
  } catch (error) {
    return "Error al guardar los datos: " + error.message;
  }
}

function generatePdfFromPreFactura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PreFactura');
  
  if (!sheet) {
    throw new Error('La hoja PreFactura no existe.');
  }
  
  // Crear una nueva hoja de cálculo temporal
  var tempSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
  var tempSheet = tempSpreadsheet.getActiveSheet();
  
  // Copiar la hoja PreFactura a la hoja temporal
  sheet.copyTo(tempSpreadsheet);
  var newSheet = tempSpreadsheet.getSheets()[1];  // La hoja copiada es la segunda hoja
  tempSpreadsheet.deleteSheet(tempSheet);  // Borrar la hoja inicial que se crea con el nuevo archivo
  newSheet.setName('PreFactura');  // Renombrar la hoja copiada
  
  // Generar el PDF
  var pdf = DriveApp.getFileById(tempSpreadsheet.getId()).getAs('application/pdf').setName('Factura.pdf');
  
  // Borrar la hoja de cálculo temporal
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  
  return pdf;
}

function getPdfUrl() {
  var pdfBlob = generatePdfFromPreFactura();
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  var contentType = pdfBlob.getContentType();
  var name = pdfBlob.getName();
  return `data:${contentType};base64,${base64Data}`;
}

function sendPdfByEmail(email) {
  var pdfFile = generatePdfFromPreFactura();
  var subject = 'Factura';
  var body = 'Adjunto encontrará la factura en formato PDF.';
  
  if (!email) {
    return "Por favor ingrese una dirección de correo válida.";
  }

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs(MimeType.PDF)]
  });
  
  return "PDF generado y enviado por correo electrónico a " + email;
}
