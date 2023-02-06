// @ts-nocheck
function onEdit() {
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMovs = archivo.getSheetByName("Movimientos JoseAntonio");
  var cellActiva = hojaMovs.getActiveCell().getValue();
  var rowActiva = hojaMovs.getActiveCell().getRow();
  var columnActiva = hojaMovs.getActiveCell().getColumn();
  var hojaResumen = archivo.getSheetByName("Resumen");
  var motivos = hojaResumen.getRange(1,4,1,2).getValues();
  
  if(rowActiva>1 && columnActiva == 2){

    var indexMot = motivos[0].indexOf(cellActiva)+4;
    if (indexMot < 4){
      hojaMovs.getActiveCell().offset(0,2).clearDataValidations();
      hojaMovs.getActiveCell().offset(0,2).clearContent();
      return;
    }
    var motivs = hojaResumen.getRange(2,indexMot,indexMot==4? 6 : 2 );
    var rvalid_motiv = SpreadsheetApp.newDataValidation().requireValueInRange(motivs);
    Logger.log(motivs.getValues());
    hojaMovs.getActiveCell().offset(0,2).setDataValidation(rvalid_motiv);
  }
  else
  if(rowActiva>1 && columnActiva == 5){
    var activis = hojaMovs.getRange(2,2,100,4).getValues();
    var suma_Efectivo=0;
    var restaEfectivo=0;
    var suma_BCP=0;
    var restaBCP=0;
    var suma_Tren=0;
    var restaTren=0;
    var suma_LPass=0;
    var restaLPass=0;
    //var conteo = 0;
    for (var i = 0; i < activis.length; i++) {
      var acty = activis[i];
      
      if(acty[0]=="Ingreso"){
        if (acty[1]=="Efectivo") suma_Efectivo += acty[3];
        else if (acty[1]=="BCP") suma_BCP += acty[3];
        else if (acty[1]=="Tren") suma_Tren += acty[3];
        else if (acty[1]=="LimaPass") suma_LPass += acty[3];
        else break;
      }
      else
      if(acty[0]=="Salida"){
        if (acty[1]=="Efectivo") restaEfectivo += acty[3];
        else if (acty[1]=="BCP") restaBCP += acty[3];
        else if (acty[1]=="Tren") restaTren += acty[3];
        else if (acty[1]=="LimaPass") restaLPass += acty[3];
        else break;
      }
      else
      if(acty[0]=="Traspaso"){
        if(acty[1]=="Efectivo" && acty[2]=="Tren"){
          restaEfectivo += acty[3];
          suma_Tren += acty[3];
        }
        else
        if(acty[1]=="Efectivo" && acty[2]=="LPass"){
          restaEfectivo += acty[3];
          suma_LPass += acty[3];
        }
        else break;
      }
      else break;
      //conteo++;
    }
    //Logger.log(conteo);
    hojaResumen.getRange(2,2).setValue(suma_Efectivo-restaEfectivo);
    hojaResumen.getRange(3,2).setValue(suma_BCP-restaBCP);
    hojaResumen.getRange(4,2).setValue(suma_Tren-restaTren);
    hojaResumen.getRange(5,2).setValue(suma_LPass-restaLPass);
    hojaResumen.getRange(9,2).setValue(suma_Efectivo);
    hojaResumen.getRange(10,2).setValue(restaEfectivo);
  }
}
