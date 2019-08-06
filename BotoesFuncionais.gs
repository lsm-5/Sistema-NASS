function BBuscaExtintor(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PaginaInicial");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
  var spreadsheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Extras");
  var spreadsheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Restrito");
  var qencontrados = 0;
  var numero = spreadsheet.getRange("A6").getValue();
  
  //setando indices corretamente
  spreadsheet2.getRange("A2").setValue(2);
  spreadsheet2.getRange("A3").setValue(3);
  var sourceRange = spreadsheet2.getRange("A2:A3");
  var destination = spreadsheet2.getRange("A2:A"+spreadsheet2.getLastRow());
  sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //fim do setando
  
  var todosExtintores = "Índice | Número | Fabricante | Tipo | Capacidade | Centro | Edificação | Pavimento\n";
  var index = -1;
  if(numero != ''){
    var formula = 'QUERY(BancoDadosGeral!A2:AF'+spreadsheet2.getLastRow()+'; "select * WHERE B= '+numero+' ";-1)';
    spreadsheet4.getRange("A2").setFormula(formula);
  }else if(numero == ''){
    var ui = SpreadsheetApp.getUi();
    ui.alert("Número do cilindro vazio.");
  }
  qencontrados = spreadsheet4.getLastRow() - 1;
  
    if (qencontrados==1){
      var info = spreadsheet4.getRange(2, 1, 1, 32).getValues();
      if(info[0][1] == ""){
        var ui = SpreadsheetApp.getUi();
        ui.alert("Extintor não encontrado.");
      }else{
        spreadsheet.getRange("A6").setValue(info[0][1]); //número
        spreadsheet.getRange("A9").setValue(info[0][2]); //fabricante
        spreadsheet.getRange("A12").setValue(info[0][3]); //tipo
        spreadsheet.getRange("A15").setValue(info[0][4]); //capacidade
        spreadsheet.getRange("D12").setValue(info[0][5]); //centro
        spreadsheet.getRange("F12").setValue(info[0][7]); //pavimento
        spreadsheet.getRange("E12").setValue(info[0][6]); //edificação
        spreadsheet.getRange("F9").setValue(info[0][8]); //carga realizada
        spreadsheet.getRange("G9").setValue(info[0][11]); //teste hidrost. realizado
        spreadsheet.getRange("D6").setValue(info[0][14]); //aspecto visual
        spreadsheet.getRange("E6").setValue(info[0][15]); //lacre
        spreadsheet.getRange("F6").setValue(info[0][16]); //bico
        spreadsheet.getRange("G6").setValue(info[0][17]); //trava
        spreadsheet.getRange("H6").setValue(info[0][18]); //circulo
        spreadsheet.getRange("I6").setValue(info[0][19]); //localização
        spreadsheet.getRange("J6").setValue(info[0][20]); //desobstruido
        spreadsheet.getRange("K6").setValue(info[0][21]); //demarcado
        spreadsheet.getRange("L6").setValue(info[0][22]); //cond do suporte
        spreadsheet.getRange("D9").setValue(info[0][23]); //tipo do suporte
        spreadsheet.getRange("E9").setValue(info[0][24]); //manometro
        spreadsheet.getRange("D15").setValue(info[0][25]); //material combust.
        spreadsheet.getRange("I9").setValue(info[0][26]); //localização extintor
        spreadsheet.getRange("A18").setValue(info[0][27]); //observações
        spreadsheet.getRange("E15").setValue(info[0][30]); //Inspenção
        spreadsheet.getRange("F15").setValue(info[0][31]); //Digitação
        spreadsheet3.getRange("A2").setValue(info[0][0]); //indice no Extra
      }
  } else if (qencontrados>1){
    for (var i = 2; i<= spreadsheet4.getLastRow(); i++){
      todosExtintores = todosExtintores + spreadsheet4.getRange("A"+i).getValue() + " " + spreadsheet4.getRange("B"+i).getValue() + " " + spreadsheet4.getRange("C"+i).getValue() + " " + spreadsheet4.getRange("D"+i).getValue() + " " + spreadsheet4.getRange("E"+i).getValue() + " " + spreadsheet4.getRange("F"+i).getValue() + " " + spreadsheet4.getRange("G"+i).getValue() + " " + spreadsheet4.getRange("H"+i).getValue() + "\n";
    }
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Foram encontrados '+qencontrados+' extintores! \n'+'Digite o índice correspondente ao seu extintor \n\n' + todosExtintores);
    if (response.getSelectedButton() == ui.Button.OK) {
      index = response.getResponseText(); //index é a resposta do usuário do indice no Banco de Dados
       for (var i = 2; i<= spreadsheet4.getLastRow(); i++){
         if(index == spreadsheet4.getRange("A"+i).getValue()){
           index = i; //index é o indice do extintor que se procura na aba Restrito
         }
       }
     var info = spreadsheet4.getRange(index, 1, 1, 32).getValues();
        spreadsheet.getRange("A6").setValue(info[0][1]); //número
        spreadsheet.getRange("A9").setValue(info[0][2]); //fabricante
        spreadsheet.getRange("A12").setValue(info[0][3]); //tipo
        spreadsheet.getRange("A15").setValue(info[0][4]); //capacidade
        spreadsheet.getRange("D12").setValue(info[0][5]); //centro
        spreadsheet.getRange("F12").setValue(info[0][7]); //pavimento
        spreadsheet.getRange("E12").setValue(info[0][6]); //edificação
        spreadsheet.getRange("F9").setValue(info[0][8]); //carga realizada
        spreadsheet.getRange("G9").setValue(info[0][11]); //teste hidrost. realizado
        spreadsheet.getRange("D6").setValue(info[0][14]); //aspecto visual
        spreadsheet.getRange("E6").setValue(info[0][15]); //lacre
        spreadsheet.getRange("F6").setValue(info[0][16]); //bico
        spreadsheet.getRange("G6").setValue(info[0][17]); //trava
        spreadsheet.getRange("H6").setValue(info[0][18]); //circulo
        spreadsheet.getRange("I6").setValue(info[0][19]); //localização
        spreadsheet.getRange("J6").setValue(info[0][20]); //desobstruido
        spreadsheet.getRange("K6").setValue(info[0][21]); //demarcado
        spreadsheet.getRange("L6").setValue(info[0][22]); //cond do suporte
        spreadsheet.getRange("D9").setValue(info[0][23]); //tipo do suporte
        spreadsheet.getRange("E9").setValue(info[0][24]); //manometro
        spreadsheet.getRange("D15").setValue(info[0][25]); //material combust.
        spreadsheet.getRange("I9").setValue(info[0][26]); //localização extintor
        spreadsheet.getRange("A18").setValue(info[0][27]); //observações
        spreadsheet.getRange("E15").setValue(info[0][30]); //Inspenção
        spreadsheet.getRange("F15").setValue(info[0][31]); //Digitação
        spreadsheet3.getRange("A2").setValue(info[0][0]); //indice no Extra
    } 
  }
}

function BLimparCampo(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PaginaInicial");
  spreadsheet.getRange("A6").setValue('');
  spreadsheet.getRange("A9").setValue('');
  spreadsheet.getRange("A12").setValue('');
  spreadsheet.getRange("A15").setValue('');
  spreadsheet.getRange("A18").setValue('');
  
  spreadsheet.getRange("D6").setValue('');
  spreadsheet.getRange("E6").setValue('');
  spreadsheet.getRange("F6").setValue('');
  spreadsheet.getRange("G6").setValue('');
  spreadsheet.getRange("H6").setValue('');
  spreadsheet.getRange("I6").setValue('');
  spreadsheet.getRange("J6").setValue('');
  spreadsheet.getRange("K6").setValue('');
  spreadsheet.getRange("L6").setValue('');
  
  spreadsheet.getRange("D9").setValue('');
  spreadsheet.getRange("E9").setValue('');
  spreadsheet.getRange("F9").setValue('');
  spreadsheet.getRange("G9").setValue('');
  
  spreadsheet.getRange("D12").setValue('');
  spreadsheet.getRange("F12").setValue('');
  spreadsheet.getRange("E12").setValue('');
  
  spreadsheet.getRange("D15").setValue('');
  spreadsheet.getRange("E15").setValue('');
  spreadsheet.getRange("F15").setValue('');
  
  spreadsheet.getRange("I9").setValue('');
}

function BSalvarAlter(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PaginaInicial");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
  var spreadsheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Extras");
  var i = spreadsheet3.getRange("A2").getValue();
  var info = spreadsheet.getRange("A1:L20").getValues();
  var infoficial = [[info[5][0], info[8][0], info[11][0], info[14][0], info[11][3], info[11][4], info[11][5], info[8][5], 0, 0, info[8][6], 0, 0, info[5][3], info[5][4], info[5][5], info[5][6], info[5][7], info[5][8], info[5][9], info[5][10], info[5][11], info[8][3], info[8][4], info[14][3], info[8][8], info[17][0], 0, 0, info[14][4], info[14][5]]];
  
  if(validacaoDados(infoficial)){ 
    //SETANDO INFORMAÇÕES
    spreadsheet2.getRange("B"+i+":AF"+i).setValues(infoficial);
    //informações da fórmulas
    var formula1 = [['=IF(I'+i+'="";"";I'+i+'+365)','=IF(J'+i+'="";"";J'+i+'-'+'TODAY()'+')']];
    var formula2 = [['=IF(L'+i+'="";"";L'+i+'+5)','=IF(M'+i+'="";"";M'+i+'-'+'YEAR(TODAY())'+')']];
    spreadsheet2.getRange("J"+i+":K"+i).setFormulas(formula1);
    spreadsheet2.getRange("M"+i+":N"+i).setFormulas(formula2);
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
    spreadsheet2.getRange("AC"+i+":AD"+i).setValues(tempo);
    //FIM DO SETANDO INFORMAÇÕES
    
    var ui = SpreadsheetApp.getUi();
    ui.alert("Alterações realizada!");
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Erro com a validação de dados.");
  }
}

function BAdicionarExtintor(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PaginaInicial");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
  var spreadsheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Extras");
  var i = spreadsheet2.getLastRow()+1;
  var info = spreadsheet.getRange("A1:L20").getValues();
  var infoficial = [[info[5][0], info[8][0], info[11][0], info[14][0], info[11][3], info[11][4], info[11][5], info[8][5], 0, 0, info[8][6], 0, 0, info[5][3], info[5][4], info[5][5], info[5][6], info[5][7], info[5][8], info[5][9], info[5][10], info[5][11], info[8][3], info[8][4], info[14][3], info[8][8], info[17][0], 0, 0, info[14][4], info[14][5]]];
  if(validacaoDados(infoficial)){
    spreadsheet2.getRange("A"+i).setValue(i);
    
    //SETANDO INFORMAÇÕES
    //informações principais
    spreadsheet2.getRange("B"+i+":AF"+i).setValues(infoficial);
    //informações da fórmulas
    var formula1 = [['=IF(I'+i+'="";"";I'+i+'+365)','=IF(J'+i+'="";"";J'+i+'-'+'TODAY()'+')']];
    var formula2 = [['=IF(L'+i+'="";"";L'+i+'+5)','=IF(M'+i+'="";"";M'+i+'-'+'YEAR(TODAY())'+')']];
    spreadsheet2.getRange("J"+i+":K"+i).setFormulas(formula1);
    spreadsheet2.getRange("M"+i+":N"+i).setFormulas(formula2);
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
    spreadsheet2.getRange("AC"+i+":AD"+i).setValues(tempo);
    //FIM DO SETANDO INFORMAÇÕES
    
    var ui = SpreadsheetApp.getUi();
    ui.alert("Extintor adicionado!");
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Erro com a validação de dados");
  }
}


function BsortAprimorado(){
   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BancoDadosGeral");
    //fase 1
    var intervalo = spreadsheet.getRange("A"+"2:AD"+spreadsheet.getLastRow()); //pega do começo ao fim do BD
    intervalo.sort({
    column: 6,
    ascending: true
    });
  
  
    //fase 2
  var index_inicial = 2;
  var index_final = 2;
  var i = 2;
  var j = 2;
  while(true){
    for (; i <= spreadsheet.getLastRow()+1; i++){
      if(spreadsheet.getRange("F"+index_inicial).getValue() == spreadsheet.getRange("F"+i).getValue() /*&& nome == spreadsheet.getRange("G"+i).getValue()*/){
        continue;
      }else{
        index_final = i-1;  
        i--;
      }
      intervalo = spreadsheet.getRange("A"+index_inicial+":AD"+index_final); //pega do começo ao fim do BD de cada lugar
      intervalo.sort({
        column: 7,
        ascending: true
      });
      
      index_inicial = index_final + 1;
      /*nome = spreadsheet.getRange("G"+index_inicial).getValue();*/
      if(index_final == spreadsheet.getLastRow()){
        break;
      }
    }
    if(index_final == spreadsheet.getLastRow()){
        break;
    }
 
  }
  
  
  //fase 3
  var index_inicial = 2;
  var index_final = 2;
  var i = 2;
  var j = 2;
  while(true){
    for (; i <= spreadsheet.getLastRow()+1; i++){
      if(spreadsheet.getRange("F"+index_inicial).getValue() == spreadsheet.getRange("F"+i).getValue() && spreadsheet.getRange("G"+index_inicial).getValue() == spreadsheet.getRange("G"+i).getValue()){
        continue;
      }else{
        index_final = i-1;  
        i--;
      }
      intervalo = spreadsheet.getRange("A"+index_inicial+":AD"+index_final); //pega do começo ao fim do BD de cada lugar
      intervalo.sort({
        column: 8,
        ascending: true
      });
      
      index_inicial = index_final + 1;
      
      if(index_final == spreadsheet.getLastRow()){
        break;
      }
    }
    if(index_final == spreadsheet.getLastRow()){
        break;
    }
 
  }
  
  //setando indices corretamente
  spreadsheet.getRange("A2").setValue(2);
  spreadsheet.getRange("A3").setValue(3);
  var sourceRange = spreadsheet.getRange("A2:A3");
  var destination = spreadsheet.getRange("A2:A"+spreadsheet.getLastRow());
  sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //fim do setando
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Banco de Dados organizado!");
}