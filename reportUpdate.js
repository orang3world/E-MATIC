/*
Notas generales: 
  Este escript trabaja sobre una hojas de calculo actualizada manualmente.
  Luego se procesa esta informacion cargando una hoja de registro.
  Está presente la posibilidad de tener una vista previa del email antes del envio
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
var sendingDate2Lines = Utilities
  .formatDate(new Date(), "GMT-3", "dd/MM/yyyy \n HH:mm:ss")

//=================================================================================
function onOpen() {
  /*
  Esta funcion habilita la aparicion de nuevos menus en la 
  interfaz grafica de la planilla. 
  */
  //-------------------------------------------------------------------------------
  SpreadsheetApp.getUi()
    .createMenu('EMAIL')
    .addItem('ACTUALIZAR', 'starting')
    .addItem('ENVIAR', 'enviar_mail')
    .addToUi();
}

//=================================================================================
function starting() {
  /*
  Esta funcion lanza dos funciones para recoger los datos y enviarlos a las hojas 
  'Informe' y la hoja de registro.
  */
  //-------------------------------------------------------------------------------
  sidebarAutoClose()
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
  var meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre']

  // Ingreso en la matriz de los encabezados
  report.push(['NOMBRE', 'APELLIDO', 'E-MAIL', 'IMPORTE', 'MENSAJE-TEXTO', 'MENSAJE-HTML', 'ROL'])

  for (let j = 0; j < roles.length; j++) { // Iteracion por las hojas con datos (en roles)

    var ssValues = dataReading(roles[j], '') // matriz con datos de la hoja dentro de la matriz roles.
    var ssHeaders = ssValues.shift() // separa los encabezados de la matriz anterior
    var lastColIndex = ssValues[0].length - 1 // indice de la ultima columna

    // DATA ITERATION
    for (let i = 0; i < ssValues.length; i++) {

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
      //var ultimo_mes = ssHeaders[lastColIndex].charAt(0).toUpperCase() + ssHeaders[lastColIndex].slice(1);
      for (let i = 0; i < meses.length; i++) {
        var re = meses[i]
        var regex = new RegExp(re, "i")
        var result = ssHeaders[lastColIndex].search(regex)
        if (result != -1) {
          var ultimo_mes = meses[i]
          break
        }
      }
      var codigo = base_html(name, surname, ultimo_mes, importe);
      var body = base_text(name, surname, ultimo_mes, importe);

      //si el email en la celda no verifica, limpiarlo y buscar comas
      var email = email.replace(/\n/g, ",") // quitar espacios
      var email = email.replace(/\s/g, ",") // quitar espacios
      var email = email.replace(/\//g, ",") // cambiar "/" por ","
      var email = email.replace(/;/g, ",") // cambiar "/" por ","
      var email = email.replace(/:/g, ",") // cambiar "/" por ","
      var email = email.replace(/,{2,}/g, ",") // cambiar varias comas juntas por ","

      while (email.search(/,/) != -1) { // busqueda de comas (detalle: busca de derecha a izq.)
        console.log('Se encontro una coma en el email')
        var subEmail = email.replace(/(.*),(.*)/, "$2")
        var email = email.replace(/(.*),(.*)/, "$1")
        console.log('subEmail : ' + subEmail)
        console.log('email restante : ' + email)

        var reportRow = [name, surname, subEmail, importe, body, codigo, rol]
        report.push(reportRow)// Carga de la fila a la matriz.

      }

      // armado de la fila de datos con las variables, para la matriz 'report'
      var reportRow = [name, surname, email, importe, body, codigo, rol]
      report.push(reportRow)// Carga de la fila a la matriz.

    }
  }
  // Si la hoja de registro no existe, la crea en el indice numero 5, como oculta.
  if (!ssAccess(systemDate, '')) {
    sp.insertSheet(systemDate, 4).hideSheet()
  } else {
    // Si la hoja de registro ya existe, la limpia de datos viejos.
    ssAccess(systemDate, '').clearContents()
  }
  // carga los datos a la hoja de registro, con la matriz 'report' ya completa
  ssAccess(systemDate, '').getRange(1, 1, report.length, report[0].length).setValues(report)
}


//=================================================================================
/*
AGREGO FUNCIONES "ENVIAR_MAIL" Y "GENERAR_INFORME".  
ADEMAS, CREE UN BASE_HTML.GS PARA TRATAR EL CONTENIDO HTML 
QUE SE ENVIA POR MAIL COMO UNA FUNCION 
SIN MOLESTAR EN EL CODIGO PRINCIPAL
*/


//=================================================================================
function generar_informe() {
  /**
   *  la funcion "generar_informe" busca generar una hoja dentro de la planilla
   *  para que puedan tener informacion de los docentes y coordinadores mas al alcance
 */
  //-------------------------------------------------------------------------------
  var meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre']
  var report = [];
  var hoja = ["Docentes", "Coordinadores"];
  report.push(['Nombre', 'Apellido', "Rol", "Mes", "Honorarios", '  Email  ', 'Actualización :  ' + sendingDate + '\nEnvio :          '])

  for (var x = 0; x < hoja.length; x++) {
    var data = dataReading(hoja[x], '');
    var ssHeaders = data.shift();
    var lastColIndex = data[0].length - 1;

    //    var ultimo_mes = ssHeaders[lastColIndex].charAt(0).toUpperCase() + ssHeaders[lastColIndex].slice(1);

    for (let i = 0; i < meses.length; i++) {
      var re = meses[i]
      var regex = new RegExp(re, "i")
      var result = ssHeaders[lastColIndex].search(regex)
      if (result != -1) {
        var ultimo_mes = meses[i]
        break
      }
    }

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
      //si el email en la celda no verifica, limpiarlo y buscar comas
      var email = email.replace(/\n/g, ",") // quitar espacios
      var email = email.replace(/\s/g, ",") // quitar espacios
      var email = email.replace(/\//g, ",") // cambiar "/" por ","
      var email = email.replace(/;/g, ",") // cambiar "/" por ","
      var email = email.replace(/:/g, ",") // cambiar "/" por ","
      var email = email.replace(/,{2,}/g, ",") // cambiar varias comas juntas por ","

      while (email.search(/,/) != -1) { // busqueda de comas (detalle: busca de derecha a izq.)
        console.log('Se encontro una coma en el email')
        var subEmail = email.replace(/(.*),(.*)/, "$2")
        var email = email.replace(/(.*),(.*)/, "$1")
        console.log('subEmail : ' + subEmail)
        console.log('email restante : ' + email)

        report.push([name, surname, hoja[x], ultimo_mes, importe, subEmail, ""]);
      }
      report.push([name, surname, hoja[x], ultimo_mes, importe, email, ""]);
    }
  }

  if (!ssAccess("Informe", "")) {
    spAccess().insertSheet("Informe");
  }
  ssAccess("Informe", "").clearContents();
  ssAccess("Informe", "").getRange(1, 1, report.length, report[0].length).setValues(report);
  var activeCell = ssAccess("Informe", "").getRange(2, 1)
  ssAccess("Informe", "").activate().setCurrentCell(activeCell)
}

