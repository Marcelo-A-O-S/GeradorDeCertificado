function SalvarModeloCertificado(modeloCertificado) {
  var sheet = SpreadsheetApp.getActive();
  var ModelosCertificadoSheet = sheet.getSheetByName("MODELOS-CERTIFICADO");
  var ultimaLinha = ModelosCertificadoSheet.getLastRow();
  var ultimaColuna = ModelosCertificadoSheet.getLastColumn();
  modeloCertificado = JSON.parse(modeloCertificado);
  if (modeloCertificado['IdModeloCertificado'] == '') {
    modeloCertificado['IdModeloCertificado'] = BuscarMaiorId("MODELOS-CERTIFICADO");
    var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    var dados = IndiceCabecalho(cabecalho, modeloCertificado);
    ModelosCertificadoSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
    return "Modelo de Certificado salvo com sucesso!"
  } else {
    for (var i = 2; i <= ultimaLinha; i++) {
      if (modeloCertificado['IdModeloCertificado'] == ModelosCertificadoSheet.getRange(i, 1).getValue()) {
        var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = IndiceCabecalho(cabecalho, modeloCertificado);
        ModelosCertificadoSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
        return "Modelo de Certificado atualizado com sucesso!"
      }
    }
  }
}
function BuscarModelosCertificado() {
  var retorno = [];
  var sheet = SpreadsheetApp.getActive();
  var ModelosCertificadoSheet = sheet.getSheetByName("MODELOS-CERTIFICADO");
  var ultimaLinha = ModelosCertificadoSheet.getLastRow();
  var ultimaColuna = ModelosCertificadoSheet.getLastColumn();
  var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  var dados = ModelosCertificadoSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
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
function BuscarModeloCertificado(IdModeloCertificado) {
  try {
    var sheet = SpreadsheetApp.getActive();
    var ModelosCertificadoSheet = sheet.getSheetByName("MODELOS-CERTIFICADO");
    var ultimaLinha = ModelosCertificadoSheet.getLastRow();
    var ultimaColuna = ModelosCertificadoSheet.getLastColumn();
    for (var i = 2; i <= ultimaLinha; i++) {
      if (ModelosCertificadoSheet.getRange(i, 1).getValue() == IdModeloCertificado) {
        var retorno = {}
        var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = ModelosCertificadoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
        cabecalho.forEach((item, index) => {
          retorno[item] = dados[index]
        })
        retorno = JSON.stringify(retorno);
        return retorno;
      }
    }
  } catch (e) {
    console.log(e);
  }
}
function DeletarModeloCertificado(IdModeloCertificado) {
  var sheet = SpreadsheetApp.getActive();
  var ModeloCertificadoSheet = sheet.getSheetByName("MODELOS-CERTIFICADO");
  var ultimaLinha = ModeloCertificadoSheet.getLastRow();
  for (var i = 2; i <= ultimaLinha; i++) {
    if (ModeloCertificadoSheet.getRange(i, 1).getValue() == IdModeloCertificado) {
      ModeloCertificadoSheet.deleteRow(i);
      return "Modelo deletado com sucesso!";
    }
  }
  return "Modelo não encontrado!";
}
function BuscarModeloCertificadoPorCurso(IdCursoTurma) {
  try {
    var sheet = SpreadsheetApp.getActive();
    var ModelosCertificadoSheet = sheet.getSheetByName("MODELOS-CERTIFICADO");
    var ultimaLinha = ModelosCertificadoSheet.getLastRow();
    var ultimaColuna = ModelosCertificadoSheet.getLastColumn();
    var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    for (var i = 2; i <= ultimaLinha; i++) {
      if (ModelosCertificadoSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCursoTurma) {
        var retorno = {}
        var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = ModelosCertificadoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
        cabecalho.forEach((item, index) => {
          retorno[item] = dados[index]
        })
        retorno = JSON.stringify(retorno);
        return retorno;
      }
    }
  } catch (e) {
    console.log(e);
  }
}
function DeletarModeloCertificadoPorCurso(IdCurso) {
  let count = 0;
  var ModelosCertificadoSheet = SpreadsheetApp.getActive().getSheetByName("MODELOS-CERTIFICADO");
  var ultimaLinha = ModelosCertificadoSheet.getLastRow();
  var ultimaColuna = ModelosCertificadoSheet.getLastColumn();
  var cabecalho = ModelosCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  for (var i = ultimaLinha; i >= 2; i--) {
    if (ModelosCertificadoSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
      ModelosCertificadoSheet.deleteRow(i);
      count++
    }
  }
  if (count > 0) {
    return `Foram deletados ${count} modelos para esse curso!`
  } else {
    return 'Nenhum modelo encontrado!'
  }
}