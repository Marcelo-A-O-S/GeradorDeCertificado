function SalvarPalestrante(palestrante) {
    var sheet = SpreadsheetApp.getActive();
    var PalestrantesSheet = sheet.getSheetByName("PALESTRANTES");
    var ultimaLinha = PalestrantesSheet.getLastRow();
    var ultimaColuna = PalestrantesSheet.getLastColumn();
    palestrante = JSON.parse(palestrante);
    if (palestrante['Id'] == '') {
        palestrante['Id'] = BuscarMaiorId("PALESTRANTES");
        var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = IndiceCabecalho(cabecalho, palestrante);
        PalestrantesSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
        return "Palestrante salvo com sucesso!"
    } else {
        for (var i = 2; i <= ultimaLinha; i++) {
            if (palestrante['Id'] == PalestrantesSheet.getRange(i, 1).getValue()) {
                var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
                var dados = IndiceCabecalho(cabecalho, palestrante);
                PalestrantesSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
                return "Palestrante atualizado com sucesso!"
            }
        }
    }
}
function BuscarPalestrante(IdPalestrante) {
    try {
        var sheet = SpreadsheetApp.getActive();
        var PalestrantesSheet = sheet.getSheetByName("PALESTRANTES");
        var ultimaLinha = PalestrantesSheet.getLastRow();
        var ultimaColuna = PalestrantesSheet.getLastColumn();
        for (var i = 2; i <= ultimaLinha; i++) {
            if (PalestrantesSheet.getRange(i, 1).getValue() == IdPalestrante) {
                var retorno = {}
                var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
                var dados = PalestrantesSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
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
function BuscarPalestrantes() {
    var retorno = [];
    var sheet = SpreadsheetApp.getActive();
    var PalestrantesSheet = sheet.getSheetByName("PALESTRANTES");
    var ultimaLinha = PalestrantesSheet.getLastRow();
    var ultimaColuna = PalestrantesSheet.getLastColumn();
    var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    var dados = PalestrantesSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
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
function DeletarPalestrante(IdPalestrante) {
    var sheet = SpreadsheetApp.getActive();
    var PalestrantesSheet = sheet.getSheetByName("PALESTRANTES");
    var ultimaLinha = PalestrantesSheet.getLastRow();
    for (var i = 2; i <= ultimaLinha; i++) {
        if (PalestrantesSheet.getRange(i, 1).getValue() == IdPalestrante) {
            PalestrantesSheet.deleteRow(i);
            return "Palestrante deletado com sucesso!";
        }
    }
    return "Palestrante não encontrado!";
}
function DeletarPalestrantesPorCurso(IdCurso) {
    let count = 0;
    var PalestrantesSheet = SpreadsheetApp.getActive().getSheetByName("PALESTRANTES");
    var ultimaLinha = PalestrantesSheet.getLastRow();
    var ultimaColuna = PalestrantesSheet.getLastColumn();
    var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    for (var i = ultimaLinha; i >= 2; i--) {
        if (PalestrantesSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
            PalestrantesSheet.deleteRow(i);
            count++
        }
    }
    if (count > 0) {
        return `Foram deletados ${count} registros de conteúdos programados para esse curso!`
    } else {
        return 'Nenhum conteúdo programático encontrado!'
    }
}
function BuscarPalestrantesPorCurso(IdCurso) {
    var retorno = [];
    var sheet = SpreadsheetApp.getActive();
    var PalestrantesSheet = sheet.getSheetByName("PALESTRANTES");
    var ultimaLinha = PalestrantesSheet.getLastRow();
    var ultimaColuna = PalestrantesSheet.getLastColumn();
    var cabecalho = PalestrantesSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    for (var i = 2; i <= ultimaLinha; i++) {
        if (PalestrantesSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
            var obj = {}
            var dados = PalestrantesSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
            cabecalho.forEach((item, index) => {
                obj[item] = dados[index]
            })
            retorno.push(obj)
        }
    }
    return retorno;
}