function base_text(name, surname, ultimo_mes, importe) {

  var template_text = '***** ADMINISTRACION FUNDACION COMPROMISO *****\n\n Hola ! :\n  ' + name + ' ' + surname + '\n\n  En esta oportunidad,\n  te solicitamos la siguiente factura:\n\n  -----------------------------------------------\n  Mes Cerrado: ' + ultimo_mes + '\n  Por un importe de pesos: $ ' + importe + '\n  En concepto de: Honorarios Profesionales\n  -----------------------------------------------\n\n  Desde ya muchas gracias.\n  Seguimos en contacto.\n\n  Atentemente :\n  Marina Ehrman.\n  Directora de administración y finanzas\n  Fundación Compromiso\n'

  return template_text;
}
