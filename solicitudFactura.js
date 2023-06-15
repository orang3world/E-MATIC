/*
Notas generales: 
  Este escript trabaja sobre una hojas de calculo actualizada manualmente.
  Luego se procesa esta informacion cargando una hoja de registro.
  Se agregan 
    una columna para el mensaje html 
    una para el mensaje version texto.
    una columna para el estado del envio 
    una columna para el error ocurrido (si este surgiera)
  Con esta hoja se arma una matriz con la que se envian los emails
  Esta presente la posibilidad de tener una vista previa del email antes del envio
  Se envian los email con confirmacion del usuario
  Se guarda la hoja como registro.

  Dentro del script:
  el nombre:
      sp significa planilla (Spreadsheet)
      ss significa hoja dentro de la planilla (Spreadsheet Sheet)
*/
// Variables para el uso de la fecha como nombre unico de las hojas creadas.

var systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd")
var sendingDate = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss")
var sendingDate2Lines = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy \n HH:mm:ss")

//=================================================================================
function onOpen() {
  /*
  Esta funcion habilita la aparicion de nuevos menus en la 
  interfaz grafica de la planilla. 
  */
  //-------------------------------------------------------------------------------
  SpreadsheetApp.getUi()
    .createMenu('EMAIL')
    .addItem('INICIAR', 'starting')
    .addItem('ENVIAR', 'enviar_mail')
    .addToUi();
}
//=================================================================================
function message(text) {
  /*
  Esta funcion permite colocar mensajes en la 
  interfaz grafica de la planilla.
  */
  //-------------------------------------------------------------------------------
  SpreadsheetApp.getActiveSpreadsheet().toast(text)
}
//=================================================================================
function messageDebugging(e) {
  /*
  Esta funcion permite colocar mensajes del error ocurrido, en la 
  interfaz grafica de la planilla y en los logs de depuracion.
  */
  //-------------------------------------------------------------------------------
  console.log('Mensaje de error : ' + e);
  message('Mensaje de error : ' + e);
}
//=================================================================================
function spAccess() {
  /*
 Esta funcion permite el acceso a la planilla donde se encuentra
  el script integrado.
 */
  //-------------------------------------------------------------------------------
  return SpreadsheetApp.getActive()
}
//=================================================================================
function ssAccess(ssName, ssIndex) {
  /*
 Esta funcion permite el acceso a una hoja de la planilla. la misma es accesible 
 tanto por su nombre, como por us numero de indice en el total de hojas. 
 Nota: un parametro estara lleno y el otro sera ´´
 */
  //-------------------------------------------------------------------------------
  var sp = spAccess()
  try {
    if (ssName != '') {
      var ss = sp.getSheetByName(ssName)
      return ss
    } else if (ssIndex != '') {
      var ss = sp.getSheets()[ssIndex]
      return ss
    } else {
      console.log('sin referencias para accesar a la ss')
      var ss = '';
      return ss
    }
  }
  catch (e) {
    messageDebugging(e)
  }
}
//=================================================================================
function dataReading(ssName, ssIndex) {
  /*
 Esta funcion permite armar una matriz con los datos dentro de 
 una de las hojas de la planilla.
 */
  //--------------------------------------------------------------------------------
  var ss = ssAccess(ssName, ssIndex)
  try {
    ss.getName()
    var ssValues = ss.getDataRange().getValues()

    var ssValuesLastRow = ssValues[ssValues.length - 1]
    if (ssValuesLastRow[0] === '' &
      ssValuesLastRow[1] === '' &
      ssValuesLastRow[2] === '' &
      ssValuesLastRow[3] != '') {

      ssValues.pop() // delete last row (subtotal)
      return ssValues

    }
  }
  catch (e) {
    messageDebugging()
  }
  return ssValues
}
//=================================================================================
function closedMonth() {
  /*
 Esta funcion devuelve el nombre del mes a facturar (mes Cerrado)
 Tiene en cuenta el numero de dia actual:
 Obtiene el MES ANTERIOR si esta antes del dia 20 incluido.
 Obtiene el MES EN CURSO para los dias posteriores al 20 de cada mes, .
 */
  //-------------------------------------------------------------------------------
  var dateToday = new Date() // Fecha completa
  var currentMonth = dateToday.getUTCMonth()  // Mes Actual (numero del 1 al 12 )
  var currentDay = dateToday.getUTCDate() // Dia del mes actual
  var months = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTRUBRE', 'NOVIEMBRE', 'DICIEMBRE']

  if (currentDay <= 20) {
    var mesCerrado = months[currentMonth - 1]
  } else {
    var mesCerrado = months(currentMonth)
  }
  //debugg console.log('mesCerrado : ' + mesCerrado + '\nmesEnCurso : ' + months[currentMonth])

  return mesCerrado
}
//=================================================================================
function starting() {
  /*
  Esta funcion lanza dos funciones para recoger los datos y enviarlos a las hojas 
  'Informe' y la hoja de registro.
  */
  //-------------------------------------------------------------------------------
  reportBuilding();
  generar_informe();
}
//=================================================================================
function reportBuilding() {
  // Esta funcion construye la hoja de registro nombrada con la fecha del sistema.
  //-------------------------------------------------------------------------------
  var report = [] // inicializa la matriz para los datos de la hoja de registro.
  var sp = spAccess() // Acceso a la planilla
  var roles = ['Docentes', 'Coordinadores']
  //var mesCerrado = closedMonth()

  //debugg console.log('ssHeaders : ' + ssHeaders)
  //debugg console.log('ssValues : ' + ssValues)


  //debugg console.log('mesCerrado : ' + mesCerrado)
  /*
    var ssTemplateValues = dataReading('Plantillas', '') // matriz con datos de la hoja 'Docentes'
    var textMessage = ssTemplateValues[1][1] // variable con la plantilla del mensaje de texto
    var htmlMessage = ssTemplateValues[2][1] // variable con la plantilla del mensaje html
  */
  // Ingreso en la matriz de los encabezados
  report.push(['NOMBRE', 'APELLIDO', 'E-MAIL', 'IMPORTE', 'MENSAJE-TEXTO', 'MENSAJE-HTML', 'ROL'])

  for (let j = 0; j < roles.length; j++) { // Iteracion por las hojas con datos (en roles)

    var ssValues = dataReading(roles[j], '') // matriz con datos de la hoja dentro de la matriz roles.
    var ssHeaders = ssValues.shift() // separa los encabezados de la matriz anterior
    var lastColIndex = ssValues[0].length - 1 // indice de la ultima columna
    /*
    var re = new RegExp(mesCerrado, "i") // patron para reconocer el mes cerrado
    var monthInToHeader = ssHeaders[lastColIndex].search(re) // devuelve -1 si falso
*/
    // DATA ITERATION
    for (let i = 0; i < ssValues.length; i++) {

      //debugg console.log('ssValues[0].length : ' + ssValues[0].length + ' , value of i = ' + i)

      var rol = roles[j]
      var name = ssValues[i][0] // variable con los nombres de la persona
      var surname = ssValues[i][1] // variable con apellido de la persona
      var email = ssValues[i][2] // variable con el email de la persona
      var importe = ssValues[i][lastColIndex].toString() // variable con el importe (ultima columna)
      var importe = new Intl.NumberFormat('es-AR',
        {
          style: 'currency',
          currency: 'ARS',
          maximumFractionDigits: 0
        })
        .format(importe)
      if (importe.charAt(0) == "$") {
        // si el caracter con el indice 0 de la variable importe es $ se le asigana a la variable importe el contenido del string contando desde el indice 1
        importe = importe.slice(1);
      }
      //var giveName = name.replace(/(^.*) (.*)$/, "$1") // variable con el primero de los nombres
      var ultimo_mes = ssHeaders[lastColIndex].charAt(0).toUpperCase() + ssHeaders[lastColIndex].slice(1);

      //debugg console.log('coincidencia re del mes cerrado en el encabezado de importes: '+monthInToHeader)
      /*
            var ccemail = '' // variable con emails para copia carbon (cc)
            var new_subject = '' // variable con el asunto del email
            var body = textMessage
            var htmlBody = htmlMessage
      */
      /* Busqueda de caracteres dentro de la plantilla de mensaje de texto
      y reemplazo por el valor de la variable de igual nombre
      */
      var codigo = base_html(name, surname, ultimo_mes, importe);
      var body = base_text(name, surname, ultimo_mes, importe);

      /*var body = body.replace('{{surname}}', surname.toUpperCase())
      .replace('{{name}}', name.toUpperCase())
      .replace('{{giveName}}', giveName)
      .replace('{{importe}}', new Intl.NumberFormat('es-AR',
        {
          style: 'currency',
          currency: 'ARS',
          maximumFractionDigits: 0
        })
        .format(importe))
      .replace('{{mesCerrado}}', mesCerrado)
*/
      /* Busqueda de caracteres dentro de la plantilla de mensaje html
      y reemplazo por el valor de la variable de igual nombre
      
      var htmlBody = htmlBody.replace('{{surname}}', surname.toUpperCase())
        .replace('{{name}}', name.toUpperCase())
        .replace('{{giveName}}', giveName)
        .replace('{{importe}}', new Intl.NumberFormat('es-AR',
          {
            style: 'currency',
            currency: 'ARS',
            maximumFractionDigits: 0
          })
          .format(importe))
        .replace('{{mesCerrado}}', mesCerrado)
*/
      //debugg console.log(body)
      console.log(body)
      //var emailVerification, sending, error
      // TEXT MESSAGE
      // armado de la fila de datos con las variables, para la matriz 'report'
      var reportRow = [name, surname, email, importe, body, codigo, rol]
      report.push(reportRow)// Carga de la fila a la matriz.

    }
  }
  // Si la hoja de registro no existe, la crea en el indice numero 5, como oculta.
  if (!ssAccess(systemDate, '')) {
    sp.insertSheet(systemDate, 5).hideSheet()
  } else {
    // Si la hoja de registro ya existe, la limpia de datos viejos.
    ssAccess(systemDate, '').clearContents()
  }
  // carga los datos a la hoja de registro, con la matriz 'report' ya completa
  ssAccess(systemDate, '').getRange(1, 1, report.length, report[0].length).setValues(report)
}

//=================================================================================
function emailPreview() {
  //Esta funcion permite una VISTA PREVIA de la version html del email.
  //-------------------------------------------------------------------------------

  try {
    if (!ssAccess(systemDate, '')) {
      message('Estos datos ya fueron enviados\nVista Previa: no disponible')
    } else {
      var ss = ssAccess(systemDate, '')
      var rowSelection = ss.getSelection().getCurrentCell().getRow()
      //console.log('Current Cell: ' + rowSelection);
      var htmlList = ss.getRange(1, 6, ss.getLastRow() - 1, 1).getValues()
      //console.log('htmlList : ' + htmlList[rowSelection])
      var html = HtmlService.createHtmlOutput("'" + htmlList[rowSelection - 1] + "'")
        .setTitle('E-MAIL : VISTA PREVIA')
      SpreadsheetApp.getUi().showSidebar(html)
    }
  }
  catch (e) {
    messageDebugging(e)
  }
}
/*
AGREGO FUNCIONES "ENVIAR_MAIL" Y "GENERAR_INFORME".  
ADEMAS, CREE UN BASE_HTML.GS PARA TRATAR EL CONTENIDO HTML 
QUE SE ENVIA POR MAIL COMO UNA FUNCION 
SIN MOLESTAR EN EL CODIGO PRINCIPAL
*/

// 
//=================================================================================
function enviar_mail() {
  /*
  la funcion "enviar_mail" primero genera la planilla "log" con todos los datos
   y luego envia cada uno de los mails
   es importante que esta funcion se ejecute una vez este generado el informe y
    chequeado que el mes sea el que corresponde
   */
  //-------------------------------------------------------------------------------
  var ui = SpreadsheetApp.getUi()

  var response = SpreadsheetApp.getUi()
    .alert('ATENCION !\n\nEsta a punto de enviar : \n\n   ' + [dataReading(systemDate, '').length - 1] + ' emails.\nEsta acción no puede ser deshecha.\n\n¿ Desea continuar ?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.NO) {
    SpreadsheetApp.getUi().alert('El envio de E-mails se ha\nCANCELADO');

    return;

  } else if (response == ui.Button.YES) {

    SpreadsheetApp.getUi().alert('El envio de E-mails ha\nCOMENZADO')

    var ssId = SpreadsheetApp.create('ssTest', 1, 1).getId()
    var ss = SpreadsheetApp.openById(ssId)

    var report = [];
    var envios = [];
    envios.push(['Envio\nÚltima Actualización:\n' + sendingDate])
    var hoja = ["Docentes", "Coordinadores"];
    report.push(['NOMBRE', 'APELLIDO', 'E-MAIL', "IMPORTE", 'ENVIO', 'MENSAJE-HTML']);

    for (var x = 0; x < hoja.length; x++) { // Iteracion por las hojas de datos 
      var data = dataReading(hoja[x], "");
      var ssHeaders = data.shift();
      var lastColIndex = data[0].length - 1;
      var ultimo_mes = ssHeaders[lastColIndex].charAt(0).toUpperCase() + ssHeaders[lastColIndex].slice(1);

      for (var i = 0; i < data.length; i++) { // Iteracion por los datos de cada hoja.
        var row = data[i];
        var name = row[0];
        var surname = row[1];
        var email = row[2];
        var importe = row[lastColIndex].toString(); // pasamos al tipo de dato string para manejar mejor la informacion
        var importe = new Intl.NumberFormat('es-AR',
          {
            style: 'currency',
            currency: 'ARS',
            maximumFractionDigits: 0
          })
          .format(importe)
        if (importe.charAt(0) == "$") {
          // si el caracter con el indice 0 de la variable importe es $ se le asigana a la variable importe el contenido del string contando desde el indice 1
          importe = importe.slice(1);
        }

        var codigo = base_html(name, surname, ultimo_mes, importe);
        var body = base_text(name, surname, ultimo_mes, importe);

        if (importe.charAt(0) !== "0") {
          // si el caracter ocn el indice 0 de la variable importe es distinto a 0 enviamos el mail
          var subject = "Facturacion de Honorarios " + ultimo_mes + ' - ' + name + ' ' + surname;

          try {
            ss.addViewer(email)
            MailApp.sendEmail(email, subject, body, { htmlBody: codigo });
            var envio = 'Exitoso'
          }
          catch (e) {
            var envio = e
          }

          report.push([name, surname, email, importe, envio, codigo]);
          envios.push([envio])

        }
      }
    }

    DriveApp.getFileById(ssId).setTrashed(true)

    if (!ssAccess(systemDate, '')) {
      // en caso de que la planilla log no exista la crea
      spAccess().insertSheet(systemDate).hideSheet();
    }
    // borramos el contenido de toda la planilla log correspondiente a la fecha actual y insertamos los datos
    ssAccess(systemDate, '').clearContents();
    ssAccess(systemDate, '').getRange(1, 1, report.length, report[0].length).setValues(report);
    ssAccess(systemDate, '').setName(sendingDate2Lines)
    ssAccess('Informe', '').getRange(1, 7, envios.length, 1).setValues(envios);
  }
  SpreadsheetApp.getUi().alert('El envio de E-mails ha\nCONCLUIDO')
}

//=================================================================================
function generar_informe() {
  /*
 la funcion "generar_informe" busca generar una hoja dentro de la planilla para que puedan tener informacion de los docentes y coordinadores mas al alcance
 */
  //-------------------------------------------------------------------------------

  var report = [];
  var hoja = ["Docentes", "Coordinadores"];
  //report.push(['Nombre', 'Apellido', 'Email', "Rol", "Mes", "Honorarios", "Envio", "Última Actualización", sendingDate])
  report.push(['Nombre', 'Apellido', "Rol", "Mes", "Honorarios", '  Email  ', "Envio\nÚltima Actualización:\n" + sendingDate])
  for (var x = 0; x < hoja.length; x++) {
    var data = dataReading(hoja[x], "");
    var ssHeaders = data.shift();
    var lastColIndex = data[0].length - 1;
    var ultimo_mes = ssHeaders[lastColIndex].charAt(0).toUpperCase() + ssHeaders[lastColIndex].slice(1);
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var name = row[0];
      var surname = row[1];
      var email = row[2];
      var importe = row[lastColIndex].toString();
      var importe = new Intl.NumberFormat('es-AR',
        {
          style: 'currency',
          currency: 'ARS',
          maximumFractionDigits: 0
        })
        .format(importe)
      if (importe.charAt(0) == "$") {
        importe = importe.slice(1);
      }
      //report.push([name, surname, email, hoja[x], ultimo_mes, importe, "", "", ""]);
      report.push([name, surname, hoja[x], ultimo_mes, importe, email, ""]);
    }
  }
  if (!ssAccess("Informe", "")) {
    spAccess().insertSheet("Informe");
  }
  ssAccess("Informe", "").clearContents();
  ssAccess("Informe", "").getRange(1, 1, report.length, report[0].length).setValues(report);
}

