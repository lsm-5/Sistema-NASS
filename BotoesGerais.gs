function BSicronizar() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Offline");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Extras");
  var spreadsheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
  var spreadsheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Restrito");
  var extintor = ""
  var qencontrados = 0;
  
  //setando indices corretamente
  spreadsheet3.getRange("A2").setValue(2);
  spreadsheet3.getRange("A3").setValue(3);
  var sourceRange = spreadsheet3.getRange("A2:A3");
  var destination = spreadsheet3.getRange("A2:A"+spreadsheet3.getLastRow());
  sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //fim do setando
  
  for(var i =2; i<= spreadsheet.getLastRow(); i++){
    
    extintor = spreadsheet.getRange(i, 1, 1, 32).getValues();
    var numero = extintor[0][1];
    var formula = 'QUERY(BancoDadosGeral!A2:AF'+spreadsheet3.getLastRow()+'; "select * WHERE B= '+numero+' ";-1)';
    spreadsheet4.getRange("A2").setFormula(formula);
    qencontrados = spreadsheet4.getLastRow();
    qencontrados--;
    
    
    if(qencontrados==1){
      if(spreadsheet4.getRange("A2").getValue() == "#N/A"){
        var indice = spreadsheet3.getLastRow()+1; 
        spreadsheet3.getRange("A"+indice).setValue(indice);
        spreadsheet3.getRange("A"+indice+":AF"+indice).setValues(extintor);
        spreadsheet3.getRange("A"+indice).setValue(indice);
        //informações da fórmulas
        var formula1 = [['=IF(I'+indice+'="";"";I'+indice+'+365)','=IF(J'+indice+'="";"";J'+indice+'-'+'TODAY()'+')']];
        var formula2 = [['=IF(L'+indice+'="";"";L'+indice+'+5)','=IF(M'+indice+'="";"";M'+indice+'-'+'YEAR(TODAY())'+')']];
        spreadsheet3.getRange("J"+indice+":K"+indice).setFormulas(formula1);
        spreadsheet3.getRange("M"+indice+":N"+indice).setFormulas(formula2);
        //informações data e hora
        var minutos = new Date().getMinutes(); //problema de zero na frente dos minutos
        if(minutos<10){
          minutos = "0"+minutos;
        }
        var hora = new Date().getHours(); //problema de zero na frente das horas
        if(hora<10){
          hora = "0"+hora;
        }
        var tempo =[[new Date(),hora+":"+minutos]];
        spreadsheet3.getRange("AC"+indice+":AD"+indice).setValues(tempo);
      }else {
        //talvez tenha q setar as formulas
        var indice = spreadsheet4.getRange("A2").getValue();
        spreadsheet3.getRange("A"+indice+":AF"+indice).setValues(extintor);
        spreadsheet3.getRange("A"+indice).setValue(indice);
        //informações da fórmulas
        var formula1 = [['=IF(I'+indice+'="";"";I'+indice+'+365)','=IF(J'+indice+'="";"";J'+indice+'-'+'TODAY()'+')']];
        var formula2 = [['=IF(L'+indice+'="";"";L'+indice+'+5)','=IF(M'+indice+'="";"";M'+indice+'-'+'YEAR(TODAY())'+')']];
        spreadsheet3.getRange("J"+indice+":K"+indice).setFormulas(formula1);
        spreadsheet3.getRange("M"+indice+":N"+indice).setFormulas(formula2);
        //informações data e hora
        var minutos = new Date().getMinutes(); //problema de zero na frente dos minutos
        if(minutos<10){
          minutos = "0"+minutos;
        }
        var hora = new Date().getHours(); //problema de zero na frente das horas
        if(hora<10){
          hora = "0"+hora;
        }
        var tempo =[[new Date(),hora+":"+minutos]];
        spreadsheet3.getRange("AC"+indice+":AD"+indice).setValues(tempo);
      }
      //limpando
      spreadsheet.getRange(i, 1, 1, 32).clear();
    }else if(qencontrados>1){
      var ui = SpreadsheetApp.getUi();
      ui.alert("foram encontrados "+qencontrados+" extintores com o número "+numero+" !");
    }
  }
  //setando indices corretamente
  spreadsheet3.getRange("A2").setValue(2);
  spreadsheet3.getRange("A3").setValue(3);
  var sourceRange = spreadsheet3.getRange("A2:A3");
  var destination = spreadsheet3.getRange("A2:A"+spreadsheet3.getLastRow());
  sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //fim do setando
}

function Binspecao(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pré-Inspenção");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
  var centro = spreadsheet.getRange("B1").getValue();
  var formula = 'QUERY(BancoDadosGeral!A2:AF'+spreadsheet2.getLastRow()+'; "select B,C,D,E,F,G,H,AA,AB WHERE F= '+"'"+centro+"'"+'";-1)';
  spreadsheet.getRange("A3").setFormula(formula);
  
}
