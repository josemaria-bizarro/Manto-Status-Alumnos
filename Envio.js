
const ss= SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1btm8_Qxyqq8mTD-M8lzWsCoUVm_q5wHSSBVtWxFqGzU/")
let SheetDATA =ss.getSheetByName("DATA");
let lsr = SheetDATA.getLastRow()+1;
let SheetDB=ss.getSheetByName("DB");

const catalogoalumnos=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY/");
let ikasleSheet =catalogoalumnos.getSheetByName("ACTIVOS_FORMATEADO");


function Notifiacion()
{ 
 
  
 var data = SheetDATA.getRange(lsr,1,1,17).getDisplayValues();
 
  var wserror = enviar(data[0][0],data[0][1],data[0][3],data[0][4],data[0][5],data[0][6],data[0][9],data[0][10],data[0][12]);
  
return wserror
}

function enviar(SOLICITANTE,DIA,MOTIVO,ALUMNO,OPCION,CORREO,PLANTEL,FILA,STATUS)
{
    var wserror =0;
    let statusPar="";
  var template = SheetDB.getRange(2,1).getValue();

  var mensaje = template.replace("{SOLICITANTE}",SOLICITANTE)
                            .replace("{DIA}",DIA)
                            .replace("{MOTIVO}",MOTIVO)
                            .replace("{ALUMNO}",ALUMNO)
                            .replace("{OPCION}",OPCION)
                            .replace("{STATUS}",STATUS)

    var bajak=0

    var fila=lsr
    
      switch (STATUS)
        {
          
          default:
            break;
          case "BAJA TEMPORAL":
            try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              AdminDirectory.Users.update(newstatus,CORREO);
              SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
              bajak=1;
            }
            catch
            {
              SheetDATA.getRange(fila,17).setValue("ERROR");
              wserror=1;
            }
            break;

          case "BAJA DEFINITIVA":
           try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              AdminDirectory.Users.update(newstatus,CORREO);
              SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
              bajak=1;
            }
            catch
            {
              SheetDATA.getRange(fila,17).setValue("ERROR");
              wserror=1;
            }
          break;

          case "BAJA ADMINISTRATIVA":
            try
            {
              var newstatus = {suspended: true,orgUnitPath:"/Suspendidos"};
              AdminDirectory.Users.update(newstatus,CORREO);
              SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
              bajak=1;
            }
            catch
            {
              SheetDATA.getRange(fila,17).setValue("ERROR");
              wserror=1;
            }
          break;

          case "CONCLUIDO":
          try
          {
            var newstatus = {suspended: true,orgUnitPath:"/Concluidos"};
            AdminDirectory.Users.update(newstatus,CORREO);
            SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
            statusPar="CONCLUIDO"
          }
          catch
          {
            SheetDATA.getRange(fila,17).setValue("ERROR");
            wserror=1;
          }
          break;

          case "SUSPENSIÓN DE AULAS":
            try
            {
               
              var newstatus = {suspended: true,orgUnitPath:"/Suspension de aulas"};
              AdminDirectory.Users.update(newstatus,CORREO);
              SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
              statusPar="INACTIVO"
              break;
            }
            catch{
              SheetDATA.getRange(fila,17).setValue("ERROR");
              wserror =1;
            }
          break;
          

          case "REACTIVACIÓN DE AULAS":
            switch (OPCION)
            {
                      case "BACH. ALIMENTOS Y BEBIDAS":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Alimentos y Bebidas";
                      break;

                      case "BACH. DISEÑO GRÁFICO":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Diseño";
                      break;
                      
                      case "BACH. COMUNICACIÓN DIGITAL":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Comunicación digital";
                      break;
                      
                      case "SAETI":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Administracion";
                      break;

                      case "LIC. ADMINISTRACIÓN":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Administracion";
                      break;

                      case "LIC. DERECHO":
                        UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Derecho";
                      break;
            };     
                try
                {
                  var newstatus = {suspended: false,orgUnitPath:UNIDAD};
                AdminDirectory.Users.update(newstatus,CORREO);
                SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
                statusPar="ACTIVO"
                }
                catch
                {
                  SheetDATA.getRange(fila,17).setValue("ERROR");
                  wserror=1;
                }
            break;

          case "ACTIVO":
              switch (OPCION)
              {
                    case "BACH. ALIMENTOS Y BEBIDAS":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Alimentos y Bebidas";
                    break;

                    case "BACH. DISEÑO GRÁFICO":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Diseño";
                    break;

                    case "BACH. DISEÑO GRÁFICO Y ARTE":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Diseño";
                    break;

                    case "SAETI":
                     UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Administracion";
                    break;

                    case "LIC. ADMINISTRACIÓN":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Administracion";
                    break;

                    case "LIC. DERECHO":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Derecho";
                    break;
               };    
               try
               {
                  var newstatus = {suspended: false,orgUnitPath:UNIDAD};
                  AdminDirectory.Users.update(newstatus,CORREO);
                  SheetDATA.getRange(fila,17).setValue("ACTUALIZADO"); 
                  statusPar="ACTIVO"; 
               }
               catch
               {
                SheetDATA.getRange(fila,17).setValue("ERROR");
                  wserror=1;
               } 
          break;

          case "RECURSADOR":
             switch (OPCION)
                {
                    case "BACH. ALIMENTOS Y BEBIDAS":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Alimentos y Bebidas";
                    break;

                    case "BACH. DISEÑO GRÁFICO":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Diseño";
                    break;

                    case "BACH. DISEÑO GRÁFICO Y ARTE":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Diseño";
                    break;

                    case "SAETI":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Bachilletratos/Administracion";
                    break;

                    case "LIC. ADMINISTRACIÓN":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Administracion";
                    break;

                    case "LIC. DERECHO":
                    UNIDAD = "/Direccion/Coordinadores/Docentes/ceec alumnos/Licenciaturas/Derecho";
                    break;
                };
        
                var newstatus = {suspended: false,orgUnitPath:UNIDAD};
                try
                {
                AdminDirectory.Users.update(newstatus,CORREO);
                SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
                statusPar="ACTIVO"
                }
                catch (err)
                {
                  SheetDATA.getRange(fila,17).setValue("ERROR");
                  wserror=1;
                  console.log(err)
                }
          break;
          
          case "EGRESADO":
            UNIDAD = "/EGRESADOS";
            var newstatus = {orgUnitPath:UNIDAD};
            try
            {
            AdminDirectory.Users.update(newstatus,CORREO);
            SheetDATA.getRange(fila,17).setValue("ACTUALIZADO");
            statusPar="EGRESADO"
            }
            catch
            {
              SheetDATA.getRange(fila,17).setValue("ERROR");
              wserror=1;
            }
          break;
       };


    if (wserror==0)
    {
        
        GmailApp.sendEmail("ceec.cafetales@ceeccafetales.edu.mx,cobranza@ceeccafetales.edu.mx,servicios.escolares@ceeccafetales.edu.mx","CAMBIO DE STATUS ACADÉMICO DEL ALUMNO "+ALUMNO, mensaje,
                                {name: 'STATUS ACADÉMICO | CEEC',noReply: true});
        /*
         GmailApp.sendEmail("ceec.cafetales@ceeccafetales.edu.mx","CAMBIO DE STATUS ACADÉMICO DEL ALUMNO "+ALUMNO, mensaje,
                                {name: 'STATUS ACADÉMICO | CEEC',noReply: true});*/

        SheetDATA.getRange(fila,16).setValue("ENVIADO");



        //OBTIENE EL REGISTRO DEL ALUMNO EN EL CATALOGO
      let ikasleDatuak =ikasleSheet.getDataRange().getDisplayValues();
      let ikasleFilter = ikasleDatuak.filter(ilara=> ilara[1]==CORREO);
      
      if (ikasleFilter.length>0)   //ENCONTRÓ AL ALUMNO EN EL CATALOGO
      {

          var wilara =ikasleFilter[0][44]   //DATO DE NUMERO DE FILA
          if (bajak==1)
            { //MODIFICA DATOS DE ALUMNOS CON BAJA
                  //obtiene el alumno a modificar en el catalogo
                var NOMBAJA =ALUMNO+"****baja"; 
                statusPar="BAJA"
                ikasleSheet.getRange(wilara,1).setValue(NOMBAJA);//MODIFICA EL NOMBRE DE ALUMNOS DADOS DE BAJA  
            }
          ikasleSheet.getRange(wilara,13).setValue(STATUS);//COLOCA EL STATUS EN EL CATALOGO DE ALUMNOS
          console.log(statusPar)
          ikasleSheet.getRange(wilara,14).setValue(statusPar);//COLOCA EL STATUS PARCIAL EN EL CATALOGO DE ALUMNOS
          ikasleSheet.getRange(wilara,40).setValue(DIA);//COLOCA LA FECHA DE ULTIMO MOVMTO EN EL CATALOGO DE ALUMNOS
      }
      else
      {
        wserror=1
          
      }
    }
   

  return wserror
}