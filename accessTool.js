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
    mDebugging(e)
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
      mDebugging()
    }
    return ssValues
  }