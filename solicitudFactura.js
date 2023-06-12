var systemDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd")
var sendingDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd \n HH:mm:ss")

//=================================================================================
function onOpen() {
  //-------------------------------------------------------------------------------
  const ui = SpreadsheetApp.getUi()

  ui
    .createMenu('EMAIL')
    .addItem('INICIAR', 'emailPreview')
    .addToUi();
}
//=================================================================================
function message(text) {
  //-------------------------------------------------------------------------------
  SpreadsheetApp.getActiveSpreadsheet().toast(text)
}
//=================================================================================
function messageDebugging(e) {
  //-------------------------------------------------------------------------------
  console.log('Mensaje de error : ' + e);
  message('Mensaje de error : ' + e);
}
//=================================================================================
function spAccess() {
  //-------------------------------------------------------------------------------
  return SpreadsheetApp.getActive()
}
//=================================================================================
function ssAccess(ssName, ssIndex) {
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
    messageDebugging()
  }
}
//=================================================================================
function dataReading(ssName, ssIndex) {
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
  //-------------------------------------------------------------------------------
  var dateToday = new Date()
  var currentMonth = dateToday.getUTCMonth()
  var currentDay = dateToday.getUTCDate()
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
  //-------------------------------------------------------------------------------
  var report = []
  var sp = spAccess()
  var ssValues = dataReading('Docentes', '')
  var ssHeaders = ssValues.shift()
  var mesCerrado = closedMonth()

  var lastColIndex = ssValues[0].length - 1
  var re = new RegExp(mesCerrado, "i")
  var monthInToHeader = ssHeaders[lastColIndex].search(re)

  //debugg console.log('ssHeaders : ' + ssHeaders)
  //debugg console.log('ssValues : ' + ssValues)


  //debugg console.log('mesCerrado : ' + mesCerrado)

  var ssTemplateValues = dataReading('Plantillas', '')
  var textMessage = ssTemplateValues[1][1]
  var htmlMessage = ssTemplateValues[2][1]

  // DATA ITERATION
  report.push(['APELLIDO', 'NOMBRE', 'E-MAIL', 'IMPORTE', 'MENSAJE-TEXTO', 'MENSAJE-HTML'])

  for (let i = 0; i < ssValues.length; i++) {

    //debugg console.log('ssValues[0].length : ' + ssValues[0].length + ' , value of i = ' + i)


    var firstName = ssValues[i][0]
    var lastName = ssValues[i][1]
    var email = ssValues[i][2]
    var amount = ssValues[i][lastColIndex] // getLastColumn
    var giveName = firstName.replace(/(^.*) (.*)$/, "$1") // only one name of firstName

    //debugg console.log('coincidencia re del mes cerrado en el encabezado de importes: '+monthInToHeader)

    var ccemail = ''
    var new_subject = ''
    var body = textMessage
    var htmlBody = htmlMessage

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

    var reportRow = [lastName, firstName, email, amount, body, htmlBody]
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