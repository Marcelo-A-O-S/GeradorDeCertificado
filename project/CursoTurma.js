function SalvarCursoTurma(cursoTurma) {
  var sheet = SpreadsheetApp.getActive();
  var CursoTurmaSheet = sheet.getSheetByName("CURSOS-TURMA");
  var ultimaLinha = CursoTurmaSheet.getLastRow();
  var ultimaColuna = CursoTurmaSheet.getLastColumn();
  cursoTurma = JSON.parse(cursoTurma);
  if (cursoTurma['IdCursoTurma'] == '') {
    cursoTurma['IdCursoTurma'] = BuscarMaiorId("CURSOS-TURMA");
    var cabecalho = CursoTurmaSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    var dados = IndiceCabecalho(cabecalho, cursoTurma);
    CursoTurmaSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
    return "Curso e turma salvo com sucesso!"
  } else {
    for (var i = 2; i <= ultimaLinha; i++) {
      if (cursoTurma['IdCursoTurma'] == CursoTurmaSheet.getRange(i, 1).getValue()) {
        var cabecalho = CursoTurmaSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = IndiceCabecalho(cabecalho, cursoTurma);
        CursoTurmaSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
        return "Curso e turma atualizado com sucesso!"
      }
    }
  }
}

function BuscarCursosTurma() {
  var retorno = [];
  var sheet = SpreadsheetApp.getActive();
  var CursoTurmaSheet = sheet.getSheetByName("CURSOS-TURMA");
  var ultimaLinha = CursoTurmaSheet.getLastRow();
  var ultimaColuna = CursoTurmaSheet.getLastColumn();
  var cabecalho = CursoTurmaSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  var dados = CursoTurmaSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
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
function BuscarCursoTurma(IdCursoTurma) {
  try {
    var sheet = SpreadsheetApp.getActive();
    var CursoTurmaSheet = sheet.getSheetByName("CURSOS-TURMA");
    var ultimaLinha = CursoTurmaSheet.getLastRow();
    var ultimaColuna = CursoTurmaSheet.getLastColumn();
    for (var i = 2; i <= ultimaLinha; i++) {
      if (CursoTurmaSheet.getRange(i, 1).getValue() == IdCursoTurma) {
        var retorno = {}
        var cabecalho = CursoTurmaSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = CursoTurmaSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
        cabecalho.forEach((item, index) => {
          retorno[item] = dados[index]
        })
        retorno = JSON.stringify(retorno);
        return retorno;
      }
    }
  }catch(e){
    console.log(e);
  }
  
}
function DeletarCursoTurma(IdCursoTurma) {
  var sheet = SpreadsheetApp.getActive();
  var CursoTurmaSheet = sheet.getSheetByName("CURSOS-TURMA");
  var ultimaLinha = CursoTurmaSheet.getLastRow();
  for (var i = 2; i <= ultimaLinha; i++) {
    if (CursoTurmaSheet.getRange(i, 1).getValue() == IdCursoTurma) {
      DeletarCertificadoPorCurso(IdCursoTurma);
      DeletarPalestrantesPorCurso(IdCursoTurma);
      DeletarConteudoProgramaticoPorCurso(IdCursoTurma);
      DeletarModeloCertificadoPorCurso(IdCursoTurma);
      CursoTurmaSheet.deleteRow(i);
      return "Curso deletado com sucesso!";
    }
  }
  return "Curso não encontrado!";
}

