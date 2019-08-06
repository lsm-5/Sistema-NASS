function validacaoDados(extintor){
  //checagem pra saber se algum campo esta em branco
  for(var i=0; i<31; i++){
    if (extintor[0][i] === ""){
      return false;
    }
  }
  return true;
}
