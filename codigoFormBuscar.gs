function OnSubmit() {
  
var form = FormApp.getActiveForm();
var formRespuestas = form.getResponses();
var formRespuesta = formRespuestas[formRespuestas.length-1];
var itemRespuestas = formRespuesta.getItemResponses();
  
var u1 = SpreadsheetApp.openById('1fGhSTEOYh1c2pZkV10FxSuad0VmHO0X8QOOunzwH13c');
var sheet = u1.getSheets();
var data = sheet[0].getDataRange().getValues();

var doc=  DocumentApp.openById(DriveApp.getFilesByName('Plantilla-MediClinicos').next().makeCopy().getId());
  doc.setName('Reporte Tipo: '+itemRespuestas[0].getResponse());
var body = doc.getBody();
  
  var table = body.appendTable();
  var row1 = table.appendTableRow();
  var col1 = row1.appendTableCell('Nombre');
  var col2 = row1.appendTableCell('Fecha Inicio');
  var col3 = row1.appendTableCell('Fecha Fin');
  var col4 = row1.appendTableCell('Tipo');
  var col5 = row1.appendTableCell('Monto');
  
  col1.setBackgroundColor('#85807F');
  col2.setBackgroundColor('#85807F');
  col3.setBackgroundColor('#85807F');
  col4.setBackgroundColor('#85807F');
  col5.setBackgroundColor('#85807F');

   for(var i = 1; i<=data.length;i++){
 
     var row = data[i];   
      
     if(row!=null && row[4]==itemRespuestas[0].getResponse()){
       
     
     var fila =  table.appendTableRow();
  
       
       
     fila.appendTableCell(row[1]);
     fila.appendTableCell(row[2]);
     fila.appendTableCell(row[3]);
     fila.appendTableCell(row[4]);
     fila.appendTableCell(row[5]);
     
    }
    
    }   
    
    doc.saveAndClose();
  var doc = DocumentApp.openById(DriveApp.getFilesByName('Reporte Tipo: '+itemRespuestas[0].getResponse()).next().getId());
  var fileDoc = DriveApp.getFileById(doc.getId());
  var pdfDoc = DriveApp.createFile(fileDoc.getAs('application/pdf'));
  pdfDoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  
  
  
  var startRow = 2; 
  var numRows = 4; 
  var sheet = u1.getActiveSheet();
  var dataRange = sheet.getRange(startRow, 2, numRows, 7);
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var correo= row[5];  
    MailApp.sendEmail(correo, 'Reporte Por Deudas Por Tipo: '+itemRespuestas[0].getResponse(), 'Reporte De Deudas De:  '+itemRespuestas[0].getResponse()+' '+pdfDoc.getUrl());
    
}



  
}
