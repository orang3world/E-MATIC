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
  function messageAlert(text) {
    /*
    Esta funcion permite colocar mensajes en la 
    interfaz grafica de la planilla.
    */
    //-------------------------------------------------------------------------------
    SpreadsheetApp.getUi().alert(text)
  }
  //=================================================================================
  function mDebugging(e) {
    /*
    Esta funcion permite colocar mensajes del error ocurrido, en la 
    interfaz grafica de la planilla y en los logs de depuracion.
    */
    //-------------------------------------------------------------------------------
    console.log('Mensaje de error : ' + e);
    message('Mensaje de error : ' + e);
  }