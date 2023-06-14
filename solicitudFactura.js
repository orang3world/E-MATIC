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
var sendingDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd \n HH:mm:ss")

//=================================================================================
function onOpen() {
  /*
  Esta funcion habilita la aparicion de nuevos menus en la 
  interfaz grafica de la planilla. 
  */
  //-------------------------------------------------------------------------------
  const ui = SpreadsheetApp.getUi()

  ui
    .createMenu('EMAIL')
    .addItem('INICIAR', 'emailPreview')
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
function reportBuilding() {
  /*
 Esta funcion construye la hoja de registro nombrada con la fecha del sistema.
 */
  //-------------------------------------------------------------------------------
  var report = [] // inicializa la matriz para los datos de la hoja de registro.
  var sp = spAccess() // Acceso a la planilla
  var ssValues = dataReading('Docentes', '') // matriz con datos de la hoja 'Docentes'
  var ssHeaders = ssValues.shift() // separa los encabezados de la matriz anterior
  var mesCerrado = closedMonth()

  var lastColIndex = ssValues[0].length - 1 // indice de la ultima columna
  var re = new RegExp(mesCerrado, "i") // patron para reconocer el mes cerrado
  var monthInToHeader = ssHeaders[lastColIndex].search(re) // devuelve -1 si falso

  //debugg console.log('ssHeaders : ' + ssHeaders)
  //debugg console.log('ssValues : ' + ssValues)


  //debugg console.log('mesCerrado : ' + mesCerrado)

  var ssTemplateValues = dataReading('Plantillas', '') // matriz con datos de la hoja 'Docentes'
  var textMessage = ssTemplateValues[1][1] // variable con la plantilla del mensaje de texto
  var htmlMessage = ssTemplateValues[2][1] // variable con la plantilla del mensaje html

  // Ingreso en la matriz de los encabezados
  report.push(['NOMBRE', 'APELLIDO', 'E-MAIL', 'IMPORTE', 'MENSAJE-TEXTO', 'MENSAJE-HTML'])

  // DATA ITERATION
  for (let i = 0; i < ssValues.length; i++) {

    //debugg console.log('ssValues[0].length : ' + ssValues[0].length + ' , value of i = ' + i)


    var firstName = ssValues[i][0] // variable con los nombres de la persona
    var lastName = ssValues[i][1] // variable con apellido de la persona
    var email = ssValues[i][2] // variable con el email de la persona
    var amount = ssValues[i][lastColIndex] // variable con el importe (ultima columna)
    var giveName = firstName.replace(/(^.*) (.*)$/, "$1") // variable con el primero de los nombres

    //debugg console.log('coincidencia re del mes cerrado en el encabezado de importes: '+monthInToHeader)

    var ccemail = '' // variable con emails para copia carbon (cc)
    var new_subject = '' // variable con el asunto del email
    var body = textMessage
    var htmlBody = htmlMessage

    /* Busqueda de caracteres dentro de la plantilla de mensaje de texto
    y reemplazo por el valor de la variable de igual nombre
    */
    var body = body.replace('{{lastName}}', lastName.toUpperCase())
      .replace('{{firstName}}', firstName.toUpperCase())
      .replace('{{giveName}}', giveName)
      .replace('{{amount}}', new Intl.NumberFormat('es-AR',
        {
          style: 'currency',
          currency: 'ARS',
          maximumFractionDigits: 0
        })
        .format(amount))
      .replace('{{mesCerrado}}', mesCerrado)

    /* Busqueda de caracteres dentro de la plantilla de mensaje html
  y reemplazo por el valor de la variable de igual nombre
  */
    var htmlBody = htmlBody.replace('{{lastName}}', lastName.toUpperCase())
      .replace('{{firstName}}', firstName.toUpperCase())
      .replace('{{giveName}}', giveName)
      .replace('{{amount}}', new Intl.NumberFormat('es-AR',
        {
          style: 'currency',
          currency: 'ARS',
          maximumFractionDigits: 0
        })
        .format(amount))
      .replace('{{mesCerrado}}', mesCerrado)

    //debugg console.log(body)
    console.log(body)
    //var emailVerification, sending, error
    // TEXT MESSAGE

    var reportRow = [firstName, lastName, email, amount, body, htmlBody]
    report.push(reportRow)

  }
  // IF NOT EXIST CREATE SSREPORT
  if (!ssAccess(systemDate, '')) {
    sp.insertSheet(systemDate, 5).hideSheet()
  } else {
    // DATA DELETE SSREPORT
    ssAccess(systemDate, '').clearContents
  }
  // DATA INPUT INTO SSREPORT
  ssAccess(systemDate, '').getRange(1, 1, report.length, report[0].length).setValues(report)
}

//=================================================================================
function emailPreview() {
  /*
 Esta funcion permite una VISTA PREVIA de la version html del email.
 */
  //-------------------------------------------------------------------------------
  try {
    var ss = ssAccess(systemDate, '')


    //ss.setActiveSelection().activateAsCurrentCell()
    var rowSelection = ss.getSelection().getCurrentCell().getRow()
    console.log('Current Cell: ' + rowSelection);

    var htmlList = ss.getRange(1, 6, ss.getLastRow() - 1, 1).getValues()

    console.log('htmlList : ' + htmlList[rowSelection])

    //var html = HtmlService.createHtmlOutputFromFile('Template').setTitle('Email preview');

    var html = HtmlService.createHtmlOutput("'" + htmlList[rowSelection - 1] + "'").setTitle('E-MAIL : VISTA PREVIA')
    //var html = HtmlService.createHtmlOutput(dLoad())
    //SpreadsheetApp.getUi().showModelessDialog(html, 'Email - Preview')
    SpreadsheetApp.getUi().showSidebar(html)
  }
  catch (e) {
    messageDebugging(e)
  }
  //SpreadsheetApp.getUi().showSidebar(html);
  //SpreadsheetApp.getUi().showModalDialog(html, 'Email - Preview')

}
//=================================================================================
function listend() {
  /*
 Esta funcion realiza el envio de los emails.
 */
  //-------------------------------------------------------------------------------
  //var body = ssTemplateValues[1][1].replace(/\\n/g, "\n")
  /*
    Object.assign(list, {
      [email]:
      {
        'to': email,
        'cc': ccemail,
        'subject': new_subject,
        'body': body,
        'htmlBody': htmlBody
      }
    })
    */
  // TIMESTAMP IN TAB OF SSREPORT
}
/*
//=================================================================================
function emailValidation(email) {
  //-------------------------------------------------------------------------------
  try {
    SpreadsheetApp.getActive().addViewer(email)
    //Exception: Invalid email: aaorange76@gmail.com
  }
  catch (e) {
    console.log('Error identification: ' + e + ' note: ' + e.getStatusCode())
  }
  SpreadsheetApp.getActive().removeViewer('aaorange75@gmail.com')

}
*/


