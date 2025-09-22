function IndiceCabecalho(cabecalho, obj){
    var dados = [];
    cabecalho.forEach(() => {
        dados.push('');
    });
    for( const [key, value] of Object.entries(obj)){
        dados[cabecalho.indexOf(key)] = value
    }
    return dados;
}
function BuscarMaiorId(nameSheet){
    var sheet = SpreadsheetApp.getActive();
    var sheetSearch = sheet.getSheetByName(nameSheet);
    var ultimaLinha = sheetSearch.getLastRow() + 1;
    var maiorNumero = 0;
    for(var i = 2; i < ultimaLinha; i++){
        if(sheetSearch.getRange(i,1).getValue() > maiorNumero){
            maiorNumero = sheetSearch.getRange(i,1).getValue()
        }
    }
    return maiorNumero +1;
}
function dataPorExtenso(data) {
  const partes = data.split("/");
  const dia = partes[0];
  const mes = parseInt(partes[1], 10);
  const ano = partes[2];
  const meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
  ];
  const dataExtenso = `${dia} de ${meses[mes - 1]} de ${ano}`;
  

  return dataExtenso;
}