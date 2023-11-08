// sript ok
function doGet()
{
  const html = HtmlService.createTemplateFromFile('Formalumnos');
  html.pubUrl ="https://script.google.com/a/macros/ceeccafetales.edu.mx/s/AKfycbzQ1WGIsHOF6nJP5wju5OHXa9R0rrCNUuRVyKxrJmo49FVL2J4TnHH3fhSEwkq3Ob4/exec"
  
  const salida = html.evaluate();
  
  return salida;
}

function include(filename) //funcion para incluir los datos de los archivos css.html y js.html funciona para ls dos indistintamente
{
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function DataAlumnos()
{
  
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY/")
  const sheetAlumnos =ss.getSheetByName('ACTIVOS_FORMATEADO');
  const dataAlumnos = sheetAlumnos.getDataRange().getDisplayValues();
  dataAlumnos.shift();
return dataAlumnos;

}


function actualizaTabla(datuak)
{
// DB SOLICITUD ALTAS/
  const ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1btm8_Qxyqq8mTD-M8lzWsCoUVm_q5wHSSBVtWxFqGzU/")
  var sheetIkasle =ss.getSheetByName("DATA");

  const datuakIkasle= sheetIkasle.getDataRange().getValues();
  var azkenIlara = sheetIkasle.getLastRow()+1;
 
  let werror=0;
  let wsalida=[]
  wsalida = [datuak]

  
  try
  {
    //obtiene rango a grabar 
     //escribe datos modificados
  sheetIkasle.getRange(azkenIlara,1,1,15).setValues(wsalida)
  ss.waitForAllDataExecutionsCompletion(2)
  
  }
  catch(err)
  {
    console.log("error al grabar datos "+err);

    werror=1;
  }

  try
  {
    werror= Notifiacion()
  
  }
  catch(err)
  {
    console.log("error al enviar correo o grabar el registro en catalogo"+err);

    werror=1;

  }


return werror;
}
