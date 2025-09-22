function SalvarConteudoProgramatico(conteudoProgramatico) {
    var sheet = SpreadsheetApp.getActive();
    var ConteudoProgramaticoSheet = sheet.getSheetByName("CONTEUDO-PROGRAMATICO");
    var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
    var ultimaColuna = ConteudoProgramaticoSheet.getLastColumn();
    conteudoProgramatico = JSON.parse(conteudoProgramatico);
    if (conteudoProgramatico['Id'] == '') {
        conteudoProgramatico['Id'] = BuscarMaiorId("CONTEUDO-PROGRAMATICO");
        var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = IndiceCabecalho(cabecalho, conteudoProgramatico);
        ConteudoProgramaticoSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
        return "Conteúdo programático salvo com sucesso!"
    } else {
        for (var i = 2; i <= ultimaLinha; i++) {
            if (conteudoProgramatico['Id'] == ConteudoProgramaticoSheet.getRange(i, 1).getValue()) {
                var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
                var dados = IndiceCabecalho(cabecalho, conteudoProgramatico);
                ConteudoProgramaticoSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
                return "Conteúdo programático atualizado com sucesso!"
            }
        }
    }
}
function BuscarConteudoProgramatico(IdConteudoProgramatico) {
    try {
        var sheet = SpreadsheetApp.getActive();
        var ConteudoProgramaticoSheet = sheet.getSheetByName("CONTEUDO-PROGRAMATICO");
        var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
        var ultimaColuna = ConteudoProgramaticoSheet.getLastColumn();
        for (var i = 2; i <= ultimaLinha; i++) {
            if (ConteudoProgramaticoSheet.getRange(i, 1).getValue() == IdConteudoProgramatico) {
                var retorno = {}
                var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
                var dados = ConteudoProgramaticoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
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
function BuscarConteudosProgramaticos() {
    var retorno = [];
    var sheet = SpreadsheetApp.getActive();
    var ConteudoProgramaticoSheet = sheet.getSheetByName("CONTEUDO-PROGRAMATICO");
    var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
    var ultimaColuna = ConteudoProgramaticoSheet.getLastColumn();
    var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    var dados = ConteudoProgramaticoSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
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
function DeletarConteudoProgramatico(IdConteudoProgramatico) {
    var sheet = SpreadsheetApp.getActive();
    var ConteudoProgramaticoSheet = sheet.getSheetByName("CONTEUDO-PROGRAMATICO");
    var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
    for (var i = 2; i <= ultimaLinha; i++) {
        if (ConteudoProgramaticoSheet.getRange(i, 1).getValue() == IdConteudoProgramatico) {
            ConteudoProgramaticoSheet.deleteRow(i);
            return "Conteúdo Programático deletado com sucesso!";
        }
    }
    return "Conteúdo Programático não encontrado!";
}
function DeletarConteudoProgramaticoPorCurso(IdCurso) {
    let count = 0;
    var ConteudoProgramaticoSheet = SpreadsheetApp.getActive().getSheetByName("CONTEUDO-PROGRAMATICO");
    var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
    var ultimaColuna = ConteudoProgramaticoSheet.getLastColumn();
    var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    for (var i = ultimaLinha; i >= 2; i--) {
        if (ConteudoProgramaticoSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
            ConteudoProgramaticoSheet.deleteRow(i);
            count++
        }
    }
    if(count > 0 ){
        return `Foram deletados ${count} registros de conteúdos programados para esse curso!`
    }else{
        return 'Nenhum conteúdo programático encontrado!'
    }
}
function BuscarConteudosProgramaticosPorCurso(IdCurso) {
    var retorno = [];
    var sheet = SpreadsheetApp.getActive();
    var ConteudoProgramaticoSheet = sheet.getSheetByName("CONTEUDO-PROGRAMATICO");
    var ultimaLinha = ConteudoProgramaticoSheet.getLastRow();
    var ultimaColuna = ConteudoProgramaticoSheet.getLastColumn();
    var cabecalho = ConteudoProgramaticoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    for (var i = 2; i <= ultimaLinha; i++) {
        if (ConteudoProgramaticoSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
            var obj = {}
            var dados = ConteudoProgramaticoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
            cabecalho.forEach((item, index) => {
                obj[item] = dados[index]
            })
            retorno.push(obj)
        }
    }
    retorno = JSON.stringify(retorno);
    return retorno
}