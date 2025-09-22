function SalvarAluno(aluno) {
  var sheet = SpreadsheetApp.getActive();
  var AlunoSheet = sheet.getSheetByName("ALUNOS");
  var ultimaLinha = AlunoSheet.getLastRow();
  var ultimaColuna = AlunoSheet.getLastColumn();
  aluno = JSON.parse(aluno);
  if(aluno['IdAluno'] == ''){
    aluno['IdAluno'] = BuscarMaiorId("ALUNOS");
    var cabecalho = AlunoSheet.getRange(1,1,1,ultimaColuna).getValues()[0];
    var dados = IndiceCabecalho(cabecalho, aluno);
    AlunoSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
    return "Aluno salvo com sucesso!"
  }else{
    for(var i = 2; i<ultimaLinha; i++){
        if(aluno['IdAluno'] == AlunoSheet.getRange(i,1).getValue()){
            var cabecalho = AlunoSheet.getRange(1,1,1,ultimaColuna).getValues()[0];
            var dados = IndiceCabecalho(cabecalho, aluno);
            AlunoSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
            return "Aluno atualizado com sucesso!"
        }
    }
  }
  
}

function BuscarAlunos(){
  var retorno = [];
  var sheet = SpreadsheetApp.getActive();
  var AlunoSheet = sheet.getSheetByName("ALUNOS");
  var ultimaLinha = AlunoSheet.getLastRow();
  var ultimaColuna = AlunoSheet.getLastColumn();
  var cabecalho = AlunoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  var dados = AlunoSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
  dados = dados.slice(1, dados.length);
  dados.forEach((dado, index) => {
    var obj = {};
    cabecalho.forEach((item, indexCabecalho) => {
      obj[item] = dado[indexCabecalho]
    })
    retorno.push(obj)
  })
  retorno = JSON.stringify(retorno);
  return retorno;
}
function BuscarAluno(IdAluno){
  var sheet = SpreadsheetApp.getActive();
  var AlunoSheet = sheet.getSheetByName("ALUNOS");
  var ultimaLinha = AlunoSheet.getLastRow();
  var ultimaColuna = AlunoSheet.getLastColumn();
  for(var i = 2; i <= ultimaLinha; i++){
    if(AlunoSheet.getRange(i, 1).getValue() == IdAluno){
        var retorno = {}
        var cabecalho = AlunoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = AlunoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
        cabecalho.forEach((item,index)=>{
            retorno[item] = dados[index]
        })
        retorno = JSON.stringify(retorno);
        return retorno;
    }
  }
}
function DeletarAluno(IdAluno){
  var sheet = SpreadsheetApp.getActive();
  var AlunoSheet = sheet.getSheetByName("ALUNOS");
  var ultimaLinha = AlunoSheet.getLastRow();
  for(var i = 2; i <= ultimaLinha; i++){
    if(AlunoSheet.getRange(i, 1).getValue() == IdAluno){ 
      AlunoSheet.deleteRow(i);
      return "Aluno deletado com sucesso!"
    }
  }
  return "Aluno não encontrado!"
}

