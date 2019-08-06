function main(){
  cabecalho(0);
  infoiniciais();
  preDados();
  dados();
  posDados();
  semiUltimasInfo();
  ultimasInfos();
}

function cabecalho(pagina) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface");
  var nome = spreadsheet.getRange("B1").getValue();
  var index = -1;
  for(var i = 1; i<= spreadsheet.getLastRow();i++){
    if(spreadsheet.getRange("B"+i).getValue() == nome){
      index = spreadsheet.getRange("A"+i).getValue();
      break;
    }
  }
  var ano = new Date().getFullYear();
  var mes = new Date().getMonth() + 1; //WHY ONE MONTH A MENOS
  var dia = new Date().getDate();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  spreadsheet.getRange('E'+((pagina*45)+2)).activate();
  spreadsheet.getCurrentCell().setValue('MEMORIAL DESCRITIVO EXTINTORES');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setFontWeight('bold')
  .setFontSize(14);
  spreadsheet.getRange('E'+((pagina*45)+3)).activate();
  spreadsheet.getCurrentCell().setValue('EXT.UFPE.SESST.'+index+'/2019');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setFontSize(14)
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('EXT.UFPE.SESST.'+index+'/2019')
  .setTextStyle(15, 23, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#ff0000')
  .build())
  .build());
  spreadsheet.getRange('E'+((pagina*45)+4)).activate();
  spreadsheet.getCurrentCell().setValue('MANUTENÇÃO');
  spreadsheet.getActiveRangeList().setFontSize(12)
  .setHorizontalAlignment('center');
  spreadsheet.getRange('G'+((pagina*45)+1)).activate();
  //spreadsheet.getCurrentCell().setValue('Página '+(pagina+1)+' de '+lastpagina)
  //.setFontSize(8);
  spreadsheet.getRange('H'+((pagina*45)+5)).activate()
  spreadsheet.getCurrentCell().setValue('Emissão:'+dia+'/'+mes+'/'+ano);
  spreadsheet.getActiveRangeList().setFontColor('#ff0000');
  spreadsheet.getRange('H'+((pagina*45)+6)).activate()
  spreadsheet.getCurrentCell().setValue('Revisão:');
  spreadsheet.getActiveRangeList().setFontColor('#ff0000');
  
  spreadsheet.getRange('A'+((pagina*45)+2)+':A'+((pagina*45)+5)).activate()
  .mergeVertically()
  .setFormula('=IMAGE("https://uploaddeimagens.com.br/images/002/226/791/full/logoUFPE.png?1564516106.png";2)');
  spreadsheet.getRange('H'+((pagina*45)+2)+':H'+((pagina*45)+4)).activate()
  .mergeVertically()
  .setFormula('=IMAGE("https://uploaddeimagens.com.br/images/002/226/803/full/Logo_SESST_colorido..png?1564516225.png";2)');
  
  spreadsheet.getRange('A'+((pagina*45)+1)+':I'+((pagina*45)+6)).activate()
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7').activate();
  spreadsheet.getActiveSheet().setFrozenRows(7);
  
};

function infoiniciais(){
  //GET INFO
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface");
  var centro = spreadsheet2.getRange("B1").getValue();
  var depart = spreadsheet2.getRange("B2").getValue();
  var coscipe = spreadsheet2.getRange("B3").getValue();
  var anexoAC = spreadsheet2.getRange("B4").getValue();
  var pavimento = spreadsheet2.getRange("B5").getValue();
  //FIM GET
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  spreadsheet.getRange('C9').activate();
  spreadsheet.getCurrentCell().setValue('UNIVERSIDADE FEDERAL DE PERNAMBUCO');
  spreadsheet.getActiveRangeList().setFontSize(12)
  .setFontWeight('bold');
  spreadsheet.getRange('C9:G9').activate()
  .mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D10').activate();
  spreadsheet.getCurrentCell().setValue('CENTRO/ÓRGÃO:');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('D11').activate();
  spreadsheet.getCurrentCell().setValue('DEPARTAMENTO:');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('D12').activate();
  spreadsheet.getCurrentCell().setValue('CLASSE DE OCUPAÇÃO COSCIPE (Art. 7°):');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('D13').activate();
  spreadsheet.getCurrentCell().setValue('CLASSE RISCO NBR 14276 (Anexos A e C):');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('D14').activate();
  spreadsheet.getCurrentCell().setValue('N° DE PAVIMENTOS:');
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('A9:I14').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  spreadsheet.getRange("E10").setValue(centro);
  spreadsheet.getRange("E11").setValue(depart);
  spreadsheet.getRange("E12").setValue(coscipe);
  spreadsheet.getRange("E13").setValue(anexoAC);
  spreadsheet.getRange("E14").setValue(pavimento);
  
};

function preDados(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  spreadsheet.getRange('B16').activate();
  spreadsheet.getCurrentCell().setValue('EXTINTORES PORTÁTEIS DE INCÊNDIO');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().setValue('1. EXTINTORES PORTÁTEIS DE INCÊNDIO');
  spreadsheet.getRange('B17').activate();
  spreadsheet.getCurrentCell().setValue('1.1 Relação dos extintores para manutenção:')
  .setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('     1.1 Relação dos extintores para manutenção:')
  .setTextStyle(0, 48, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());
  spreadsheet.getRange('A19').activate();
  spreadsheet.getCurrentCell().setValue('N° Cilindro');
  spreadsheet.getRange('B19').activate();
  spreadsheet.getCurrentCell().setValue('Tipo');
  spreadsheet.getRange('C19').activate();
  spreadsheet.getCurrentCell().setValue('Capacidade');
  spreadsheet.getRange('D19').activate();
  spreadsheet.getCurrentCell().setValue('Localização');
  spreadsheet.getRange('D19:E19').activate()
  .mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('F19').activate();
  spreadsheet.getCurrentCell().setValue('Edificação');
  spreadsheet.getRange('G19').activate();
  spreadsheet.getCurrentCell().setValue('Pavimento');
  spreadsheet.getRange('H19').activate();
  spreadsheet.getCurrentCell().setValue('Teste Hidrostático');
  spreadsheet.getRange('A19:H19').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  .setBackground('#ff0000')
  .setFontColor('#ffffff')
  .setFontWeight('bold');
}

function dados(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BDmemorial");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Informações");
  var spreadsheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  
  var nome = spreadsheet2.getRange("c2").getValue();
  var formula = 'QUERY(BDImportado!A2:AF'+spreadsheet2.getLastRow()+'; "select B,D,E,AA,G,H,N WHERE F= '+"'"+nome+"'"+' ";-1)';
  spreadsheet.getRange("A2").setFormula(formula);
  
  var Qextintores = spreadsheet.getLastRow();
  Qextintores--;
  var indexinicial = 20;
  for(var i = 1; i<= Qextintores; i++){
    var extintor = spreadsheet.getRange(i+1,1,1,7).getValues();
    spreadsheet3.getRange('D'+indexinicial+':E'+indexinicial).activate()
    .mergeAcross();
    spreadsheet3.getRange("A"+indexinicial).setValue(extintor[0][0]);
    spreadsheet3.getRange("B"+indexinicial).setValue(extintor[0][1]);
    spreadsheet3.getRange("C"+indexinicial).setValue(extintor[0][2]);
    spreadsheet3.getRange("D"+indexinicial).setValue(extintor[0][3]);
    spreadsheet3.getRange("D"+indexinicial+":E"+indexinicial).activate()
    spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    spreadsheet3.getRange("F"+indexinicial).setValue(extintor[0][4]);
    spreadsheet3.getRange("G"+indexinicial).setValue(extintor[0][5]);
    spreadsheet3.getRange("H"+indexinicial).setValue(extintor[0][6]);
    indexinicial++;
  } 
}

function posDados(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Informações");

  var ultima = spreadsheet.getLastRow() + 2;
  var informacoes = spreadsheet2.getRange(2, 5, 17, 2).getValues();
  var texto = "Nº total de extintores = "+informacoes[16][0]+", sendo: "+informacoes[0][0]+" de AP de 10L + "+informacoes[1][0]+" de AP de 75L + "+informacoes[2][0]+" de PQS (ABC) de 4Kg + "+informacoes[3][0]+" de PQS (ABC) de 6Kg + "+informacoes[4][0]+" de PQS (BC) de 4Kg + "+informacoes[5][0]+" PQS (BC) de 6Kg + "+informacoes[6][0]+" PQS (BC) de 8Kg + "+informacoes[7][0]+" PQS (BC) de 10Kg  + "+informacoes[8][0]+" PQS (BC) de 12Kg + "+informacoes[9][0]+" PQS (BC) de 20Kg + "+informacoes[10][0]+" PQS (BC) de 50Kg + "+informacoes[11][0]+" de CO2 de 6Kg + "+informacoes[12][0]+" de CO2 de 8Kg + "+informacoes[13][0]+" de CO2 de 10Kg + "+informacoes[14][0]+" de CO2 de 12Kg + "+informacoes[15][0]+" de CO2 de 25Kg";
  
  spreadsheet.getRange("B"+ultima).setValue(texto);
  spreadsheet.getRange('B'+ultima+':H'+(ultima+3)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  var texto1 = "LEGENDA: PQS - Pó Químico Seco; CO2 – Dióxido de carbono; C – Conforme; NC – Não Conforme; NA – Não Aplicável."
  spreadsheet.getRange("B"+(spreadsheet.getLastRow()+4)).setValue(texto1);
  spreadsheet.getRange("B"+(spreadsheet.getLastRow())+":H"+(spreadsheet.getLastRow()+1)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function semiUltimasInfo(){
  //http://uploaddeimagens.com.br/images/002/229/372/full/IMAGENS_PARA_MEMORIAL.jpg?1564597914
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  var ultimo = spreadsheet.getLastRow() + 1;
  spreadsheet.getRange('B'+ultimo).activate();
  spreadsheet.getCurrentCell().setValue('1.2 Instruções Gerais');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('B'+(ultimo+1)).activate();
  spreadsheet.getCurrentCell().setValue('     1.2.1 Especificação da instalação');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('C'+(ultimo+2)).activate();
  spreadsheet.getCurrentCell().setValue('Instalar o círculo ou seta indicativa acima do extinto;');
  spreadsheet.getRange('C'+(ultimo+3)).activate();
  spreadsheet.getCurrentCell().setValue('Colocar o extintor no suporte apropriado, piso ou parede (Fig. 1 e 2);');
  spreadsheet.getRange('C'+(ultimo+4)).activate();
  spreadsheet.getCurrentCell().setValue('Os extintores não devem ter a sua parte superior acima de 1,60 m do piso, quando instalados na parede;');
  spreadsheet.getRange('C'+(ultimo+4)+':H'+(ultimo+5)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('C'+(ultimo+6)).activate();
  spreadsheet.getCurrentCell().setValue('Os extintores devem ser instalados em locais onde:');
  spreadsheet.getRange('C'+(ultimo+7)).activate();
  spreadsheet.getCurrentCell().setValue('     a) Haja menor probabilidade de o fogo bloquear seu acesso;');
  spreadsheet.getRange('C'+(ultimo+8)).activate();
  spreadsheet.getCurrentCell().setValue('     b) Sejam visíveis;');
  spreadsheet.getRange('C'+(ultimo+9)).activate();
  spreadsheet.getCurrentCell().setValue('      c) Conservem-se protegidos contra golpes e intempéries.');
  spreadsheet.getRange('C'+(ultimo+10)).activate();
  spreadsheet.getCurrentCell().setValue('     d) Não fiquem encobertos ou obstruídos;');
  spreadsheet.getRange('C'+(ultimo+11)).activate();
  spreadsheet.getCurrentCell().setValue('Fazer a demarcação de 1m² de isolamento com a fita adesiva (Fig. 3) e manter esta área desobstruída;');
  spreadsheet.getRange('C'+(ultimo+11)+':H'+(ultimo+12)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('C'+(ultimo+13)).activate();
  spreadsheet.getCurrentCell().setValue('Comunicar à empresa especializada em manutenção de extintores qualquer irregularidade/uso indevido/danos.');
  spreadsheet.getRange('C'+(ultimo+13)+':H'+(ultimo+14)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getCurrentCell().setValue('Comunicar à empresa especializada em manutenção de extintores qualquer irregularidade/ uso indevido/ danos.');
  
  spreadsheet.getRange('B'+(ultimo+15)+':H'+(ultimo+31)).activate()
  .merge()
  .setFormula('=IMAGE("http://uploaddeimagens.com.br/images/002/229/372/full/IMAGENS_PARA_MEMORIAL.jpg?1564597914";2)');
}

function ultimasInfos(){
  var spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface");
  var depart = spreadsheet2.getRange("B2").getValue();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  var ultimo = spreadsheet.getLastRow() + 18;
  spreadsheet.getRange('B'+ultimo).activate();
  spreadsheet.getCurrentCell().setValue('     1.2.2. Plano de manutenção');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('B'+(ultimo+1)).activate();
  spreadsheet.getCurrentCell().setValue('Intervenção visual');
  spreadsheet.getRange('B'+(ultimo+1)+':D'+(ultimo+1)).activate()
  .mergeAcross();
  
  spreadsheet.getRange('E'+(ultimo+1)).activate();
  spreadsheet.getCurrentCell().setValue('Periodicidade');
  spreadsheet.getRange('F'+(ultimo+1)).activate();
  spreadsheet.getCurrentCell().setValue('Responsabilidade');
  spreadsheet.getRange('F'+(ultimo+1)+':H'+(ultimo+1)).activate()
  .mergeAcross();
  spreadsheet.getRange('B'+(ultimo+1)+':H'+(ultimo+1)).activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000')
  .setFontColor('#ffffff')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('B'+(ultimo+2)).activate();
  spreadsheet.getCurrentCell().setValue('Inspeção visual');
  spreadsheet.getRange('B'+(ultimo+3)).activate();
  spreadsheet.getCurrentCell().setValue('Recarga');
  spreadsheet.getRange('B'+(ultimo+4)).activate();
  spreadsheet.getCurrentCell().setValue('Teste hidrostático');
  spreadsheet.getRange('E'+(ultimo+2)).activate();
  spreadsheet.getCurrentCell().setValue('Semestral');
  spreadsheet.getRange('E'+(ultimo+3)).activate();
  spreadsheet.getCurrentCell().setValue('Anual');
  spreadsheet.getRange('E'+(ultimo+4)).activate();
  spreadsheet.getCurrentCell().setValue('5 anos');
  spreadsheet.getRange('F'+(ultimo+2)).activate();
  spreadsheet.getCurrentCell().setValue('SESST');
  spreadsheet.getRange('F'+(ultimo+3)).activate();
  spreadsheet.getCurrentCell().setValue('Empresa especializada em manutenção de extintores credenciada pelo Corpo de Bombeiros');
  spreadsheet.getRange('F'+(ultimo+3)+':H'+(ultimo+4)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('B'+(ultimo+2)+':D'+(ultimo+2)).activate()
  .mergeAcross();
  spreadsheet.getRange('B'+(ultimo+3)+':D'+(ultimo+3)).activate()
  .mergeAcross();
  spreadsheet.getRange('B'+(ultimo+4)+':D'+(ultimo+4)).activate()
  .mergeAcross();
  spreadsheet.getRange('F'+(ultimo+2)+':H'+(ultimo+2)).activate()
  .mergeAcross();
  spreadsheet.getRange('B'+(ultimo+2)+':H'+(ultimo+2)).activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('B'+(ultimo+6)).activate();
  spreadsheet.getCurrentCell().setValue('2. FUNDAMENTAÇÃO LEGAL');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('B'+(ultimo+7)).activate();
  spreadsheet.getCurrentCell().setValue('          Norma Regulamentadora 23 do Ministério do Trabalho – Proteção contra incêndio COSCIPE - Código de Segurança contra Incêndio e Pânico para o Estado de Pernambuco do Corpo de Bombeiros de Pernambuco.');
  spreadsheet.getRange('B'+(ultimo+7)+':H'+(ultimo+9)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('B'+(ultimo+11)).activate();
  spreadsheet.getCurrentCell().setValue('3. CONSIDERAÇÕES IMPORTANTES');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('B'+(ultimo+12)).activate();
  spreadsheet.getCurrentCell().setValue('          Este memorial descritivo se limita a relacionar os extintores de incêndio do(a) '+ depart +', que devem passar pelo processo de manutenção.');
  spreadsheet.getRange('B'+(ultimo+12)+':H'+(ultimo+14)).activate()
  .merge();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('G'+(ultimo+17)).activate(); 
  
}

function limpar(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicial");
  var ultimo = spreadsheet.getLastRow();
  spreadsheet.getRange("A1"+":I"+(ultimo)).clear();
   var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7').activate();
  spreadsheet.getActiveSheet().setFrozenRows(0);
}