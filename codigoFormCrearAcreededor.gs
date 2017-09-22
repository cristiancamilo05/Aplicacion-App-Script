var form = FormApp.getActiveForm();
 var formRespuestas = form.getResponses();
 var formRespuesta = formRespuestas[formRespuestas.length-1];
 var itemRespuestas = formRespuesta.getItemResponses();
  
  
  var calendario  = CalendarApp.getCalendarById('odqpk2j0glmf2i8p4kfl282vfg@group.calendar.google.com');
  calendario.createEventSeries(itemRespuestas[3].getResponse(), new Date(itemRespuestas[1].getResponse()), new Date(itemRespuestas[2].getResponse()), CalendarApp.newRecurrence().addMonthlyRule(), {description: 'Deuda abquirida con: '+itemRespuestas[0].getResponse()+ 'Monto a Cancelar: '+itemRespuestas[4].getResponse() ,guests: 'camiloarcila05@gmail.com, pablo.a.s.mendoza@gmail.com', sendInvites:true});
  
}
