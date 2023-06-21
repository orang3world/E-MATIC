//=================================================================================
function sidebarAutoClose() {
  //Esta funcion permite cerrar la  VISTA PREVIA de la version html del email.
  //-------------------------------------------------------------------------------
  var p = PropertiesService.getScriptProperties();

  if (p.getProperty("sidebar") == "open") {
    var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
    SpreadsheetApp.getUi().showSidebar(html);

    p.setProperty("sidebar", "close");
  }

}
//=================================================================================
function emailPreview() {
  //Esta funcion permite una VISTA PREVIA de la version html del email.
  //-------------------------------------------------------------------------------

  try {
    if (!ssAccess(systemDate, '')) {
      messageAlert('Estos datos ya fueron enviados\nVista Previa: no disponible')
    } else {
      var p = PropertiesService.getScriptProperties();
      p.setProperty("sidebar", "open");

      var ss = ssAccess(systemDate, '')
      var rowSelection = ss.getSelection().getCurrentCell().getRow()
      //console.log('Current Cell: ' + rowSelection);
      var htmlList = ss.getRange(1, 6, ss.getLastRow(), 1).getValues()
      //console.log('htmlList : ' + htmlList[rowSelection])
      var html = HtmlService.createHtmlOutput("'" + htmlList[rowSelection - 1] + "'")
        .setTitle('E-MAIL : VISTA PREVIA')
      SpreadsheetApp.getUi().showSidebar(html)

      // Creates a trigger that runs 45 seconds later
    }
  }
  catch (e) {
    mDebugging(e)
  }
}