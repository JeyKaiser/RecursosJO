/* crea un menu con 2 items en la barra de herramientas, que llama a un archivo html:
   1: Devuelve valores y son procesados y retorna los resultados a Html
      copia de seguridad: D:\Apps Script\1- CALCULADORA SESGOS GOOGLE SHEETS
   2: Toma el dato ingresado y lo busca en otro libro y devuelve 3 valores 
      y los muestra en una tabla
*/

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Crea un menú en la barra superior del documento
    ui.createMenu('RecursosJO')
        .addItem('Calcular Sesgos', 'mostrarCalculadora')
        .addItem('Saldo Inventario', 'buscadorItem')
        .addToUi();
  }
  
  function mostrarCalculadora() {
    var html = HtmlService.createHtmlOutputFromFile('Calculadora')
        .setTitle('Calculadora de Sesgos')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function buscadorItem() {
    var html = HtmlService.createHtmlOutputFromFile('BuscadorInventario')
        .setTitle('Buscar Inventario')
        .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);}
  
  //---------------------------------------------------------------------------------------
  
  // Función llamada desde la barra lateral
  function procesarValores(valor1, valor2, valor3, valor4) {
    // Convertir valores a números
    var num1 = parseFloat(valor1);
    var num2 = parseFloat(valor2);
    var num3 = parseFloat(valor3);
    var num4 = parseFloat(valor4);
  
    // Calcular al sesgo, a travez y al hilo  
    var sesgo = Math.round((num1) * (num3 * 0.01) / (num2 - (num2 * (num4/100)))* 100) / 100;
    var travez = Math.round((num1 * (num3 * 0.01)) / num2* 100) / 100;
    var hilo = Math.round((num1 * (num3 * 0.01)) / num2* 100) / 100;     
  
    // Obtener la hoja de cálculo activa y la celda activa
    //var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //var celdaActiva = hoja.getActiveCell();
  
    // Devolver los resultados para mostrar en la interfaz de la aplicación
    return {
      sesgo: sesgo,
      travez: travez,
      hilo: hilo
    };
  }
  
  function buscarEnInventario(codigoSap) {
    try {
      // Obtener la hoja activa
      var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      
      // Establecer el ID del otro libro y el nombre de la hoja objetivo
      var targetBookId = "1xnV6wMtkbbORb8H5IKfA6iih3h4hwtkRsP581gzVkTw";
      var targetSheetName = "CONTEO 14-12-23"; 
      
      // Obtener el libro y verificar si existe
      var targetBook = SpreadsheetApp.openById(targetBookId);
      if (!targetBook) {
        throw new Error("No se pudo abrir el otro libro.");
      }
      
      // Obtener la hoja objetivo y verificar si existe
      var targetSheet = targetBook.getSheetByName(targetSheetName);
      if (!targetSheet) {
        throw new Error("No se encontró la hoja en el otro libro.");
      }
      
      // Obtener los datos de la hoja activa
      var data = activeSheet.getDataRange().getValues();
      
      // Obtener los datos de la hoja objetivo
      var targetData = targetSheet.getRange("A:F").getValues();
      
      // Buscar el término en la columna A del libro objetivo
      var found = false;
      var rowData;
      for (var i = 0; i < targetData.length; i++) {
        if (targetData[i][0] === codigoSap) {
          found = true;
          rowData = targetData[i];
          break;
        }
      }
      
      // Mostrar mensaje según si se encontró o no el término
      if (found) {
        // Devolver los valores de las columnas C y F de la misma fila
        return {
          Descripcion: rowData[1], // Índice 2 representa la columna B
          Saldo      : rowData[4], // Índice 2 representa la columna C
          Ubicacion  : rowData[5]  // Índice 5 representa la columna F
        };
      } else {
        throw new Error('Tela no encontrada en el inventario.');
      }
    } catch (error) {
      // Manejar errores
      throw new Error(error.message);
    }
  }
  
  
  
  
  