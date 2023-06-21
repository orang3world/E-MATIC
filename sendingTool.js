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

  try {

    if (!ssAccess(systemDate, '')) {
      messageAlert('Estos datos ya fueron enviados y archivados:\n Actualice para un nuevo Envio')

      sidebarAutoClose()

    } else {

      sidebarAutoClose()


      //-------------------------------------------------------------------------------
      var sessionEmail = Session.getActiveUser().getEmail();

      var responseMode = SpreadsheetApp
        .getUi()
        .alert('ATENCION !!\n\nMODO DE PRUEBA\n\nEn este modo, todos los E-mails\nse enviaran a una UNICA dirección:\n\n   ' + sessionEmail + '\n\n¿ Desea continuar en este modo?', ui.ButtonSet.YES_NO);

      if (responseMode == ui.Button.NO) {
        var sessionMode = 'REAL'

      } else if (responseMode == ui.Button.YES) {
        var sessionMode = 'PRUEBA'

      }
      //-------------------------------------------------------------------------------
      var response = SpreadsheetApp
        .getUi()
        .alert('CONFIRMACION !!\nMODO DE ENVIO: ___ ' + sessionMode + '\n\nEsta acción no puede ser deshecha.\n\nEsta a punto de enviar : \n\n   ' + [dataReading(systemDate, '').length - 1] + ' ___ E-mails.\n\n¿ Desea continuar ?', ui.ButtonSet.YES_NO);

      if (response == ui.Button.NO) {
        message('El envio de E-mails se ha\nCANCELADO');

        return;

      } else if (response == ui.Button.YES) {

        message('El envio de E-mails ha\nCOMENZADO')

        // creacion de planilla para verificacion de direcciones de email.
        var spTestId = SpreadsheetApp.create('ssTest', 1, 1).getId()
        var spEmailTest = SpreadsheetApp.openById(spTestId)

        var report = [];
        var envios = [];
        var hoja = ["Docentes", "Coordinadores"];
        var meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre']
        report.push(['NOMBRE', 'APELLIDO', 'E-MAIL', "IMPORTE", 'MENSAJE-TEXTO', 'ENVIO', 'MENSAJE-HTML']);
        // report.push([name, surname, importe, body, envio, email, codigo]);

        for (var x = 0; x < hoja.length; x++) { // Iteracion por las hojas de datos 
          var data = dataReading(hoja[x], "");
          var ssHeaders = data.shift();
          var lastColIndex = data[0].length - 1;
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

                if (sessionMode == 'PRUEBA') {
                  var email = sessionEmail
                  spEmailTest.addViewer(email) // linea de verificacion del email, si falla , no envia.

                  MailApp
                    .sendEmail(email, subject, body, { htmlBody: codigo });

                } else {

                  spEmailTest.addViewer(email) // linea de verificacion del email, si falla , no envia.
                  // La siguiente linea envia emails en el modo REAL, usa las direcciones reales.
                  //MailApp.sendEmail(email, subject, body, { htmlBody: codigo });
                }

                var envio = 'Exitoso'
              }
              catch (e) {
                var envio = e
              }

              report.push([name, surname, email, importe, body, envio, codigo]);
              envios.push([envio])

            }
          }
        }
        var timeStamp = ssAccess('Informe', '').getRange(1, 7).getValue().toString()

        var timeStamp = timeStamp + Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss")
        envios.unshift([timeStamp])

        // eliminacion de planilla para verificacion de direcciones de email.
        DriveApp.getFileById(spTestId).setTrashed(true)

        if (!ssAccess(systemDate, '')) {
          // en caso de que la planilla log no exista la crea
          spAccess().insertSheet(systemDate).hideSheet();
        }

        var docentesLog = ssAccess('Docentes', '').getDataRange().getValues();
        var coordinadoresLog = ssAccess('Coordinadores', '').getDataRange().getValues();

        // borramos el contenido de toda la planilla log correspondiente a la fecha actual y insertamos los datos

        ssAccess(systemDate, '').clearContents();

        ssAccess(systemDate, '').getRange(1, 1, docentesLog.length, docentesLog[0].length).setValues(docentesLog);
        ssAccess(systemDate, '').getRange(1, docentesLog[0].length + 2, coordinadoresLog.length, coordinadoresLog[0].length).setValues(coordinadoresLog);
        ssAccess(systemDate, '').getRange(docentesLog.length + 2, 1, report.length, report[0].length).setValues(report);

        ssAccess(systemDate, '').setName(sendingDate2Lines)

        ssAccess('Informe', '').getRange(1, 7, envios.length, 1).setValues(envios);
      }
      message('El envio de E-mails ha\nCONCLUIDO')
    }
  }
  catch (e) {
    message(e)
  }
}