function GerarCertificado(idCertificado) {
  try {
    var certificado = BuscarCertificado(idCertificado);
    certificado = JSON.parse(certificado);
    var modelo = BuscarModeloCertificado(certificado.IdModeloCertificado);
    modelo = JSON.parse(modelo);
    var aluno = BuscarAluno(certificado.IdAluno);
    aluno = JSON.parse(aluno);
    var conteudosProgramaticos = BuscarConteudosProgramaticosPorCurso(certificado.IdCursoTurma);
    conteudosProgramaticos = JSON.parse(conteudosProgramaticos);
    certificado.ChCurso += conteudosProgramaticos.reduce((total, conteudoProgramatico) => {
      return total + parseInt(conteudoProgramatico.CargaHoraria, 10);
    }, 0);
    var palestrantes = BuscarPalestrantesPorCurso(certificado.IdCursoTurma)
    var folder = DriveApp.getFolderById(modelo.IdPasta);
    var copia = DriveApp.getFileById(modelo.IdModelo).makeCopy(certificado.Nome + " - Certificado " + new Date().toLocaleDateString('pt-BR'), folder);
    certificado.Codigo = Utilities.getUuid();
    certificado.Certificado = "https://drive.google.com/uc?export=view&id=" + copia.getId();
    const doc = DocumentApp.openById(copia.getId());
    const docId = doc.getId();
    var qrUrl = "https://quickchart.io/qr?text=" + encodeURIComponent(certificado.Certificado) + "&size=60";
    var response = UrlFetchApp.fetch(qrUrl);
    var imageBlob = response.getBlob().setName("qrcode.png");
    const file = DriveApp.getFileById(doc.getId());
    folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
    const body = doc.getBody();
    var footer = doc.getFooter() || doc.addFooter(); footer.clear();
    var paragraph = footer.appendParagraph("");
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    paragraph.appendInlineImage(imageBlob).setWidth(60).setHeight(60);
    var paragraphCodigo = footer.appendParagraph(`${certificado.Codigo}`)
    paragraphCodigo.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    paragraphCodigo.editAsText().setBold(true).setFontSize(10);
    // --- Título --- 
    var paragraphTitle = body.appendParagraph("CERTIFICADO")
      .setHeading(DocumentApp.ParagraphHeading.TITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    paragraphTitle.editAsText()
      .setFontFamily("Algerian")
      .setFontSize("48")
      .setBold(true)
      .setUnderline(true);
    // --- Texto principal --- 
    let DtInicio = new Date(certificado.DtInicio).toLocaleDateString('pt-BR');
    let DtFim = new Date(certificado.DtFim).toLocaleDateString('pt-BR');
    let texto = `Lar Projetos Elétricos Inscrita no CNPJ 23.361.953/0001-80 Certifica que Realizou no período de ${dataPorExtenso(DtInicio)} à ${dataPorExtenso(DtFim)}, Curso teórico-prático de ${certificado.Curso}, com duração de ${certificado.ChCurso} Horas, do qual participou ${aluno.Genero == "Masculino" ? "o Sr." : "a Sra."} ${certificado.Nome}, CPF ${certificado.Cpf}. `;
    var ParagraphcontentData = body.appendParagraph(texto)
      .setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
      .setFontSize("20");
    ParagraphcontentData.editAsText()
      .setFontFamily("Oswald");
    // --- Local e data --- 
    var localData = body.appendParagraph(`Bocaiúva, ${dataPorExtenso(new Date().toLocaleDateString('pt-BR'))}`)
      .setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    localData.editAsText()
      .setFontSize("16")
      .setFontFamily("Oswald");
    // --- Assinaturas --- 
    body.appendParagraph("--------------------------------------------")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var assinaturaAluno = body.appendParagraph(`${certificado.Nome}`)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    assinaturaAluno.editAsText()
      .setFontSize("16")
      .setFontFamily("Oswald")
      .setBold(true);
    var tabelasPalestrantes = body.appendTable()
    const linha = tabelasPalestrantes.appendTableRow();
    palestrantes.forEach(item => {
      const cell = linha.appendTableCell(
        "--------------------------------\n" +
        item.Nome + "\n" +
        item.Formacao + "\n" +
        "Instrutor/Palestrante"
      );
      for (let i = 0; i < cell.getNumChildren(); i++) {
        const child = cell.getChild(i);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          child.asParagraph()
            .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
            .editAsText()
            .setFontFamily("Oswald")
            .setFontSize(16);
        }
      }
    })
    tabelasPalestrantes.setBorderWidth(0);
    body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
    const linhas = conteudosProgramaticos.map(c => [
      c.Conteudo,
      c.CargaHoraria + " Horas"
    ]);
    const tabelaConteudo = body.appendTable([
      ["Conteúdo Programático", "Horas/Aulas"],
      ...linhas
    ])
    tabelaConteudo.getRow(0).editAsText()
      .setBold(true)
      .setFontSize(14);

    for (let i = 0; i < tabelaConteudo.getNumRows(); i++) {
      const row = tabelaConteudo.getRow(i);
      for (let j = 0; j < row.getNumCells(); j++) {
        const cell = row.getCell(j);

        for (let k = 0; k < cell.getNumChildren(); k++) {
          const child = cell.getChild(k);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            child.asParagraph().editAsText().setFontSize(14);
          }
        }
      }
    }
    tabelaConteudo.setBorderWidth(1);
    body.appendParagraph("");
    body.appendParagraph("--------------------------------------------")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var assinaturaDiretor = body.appendParagraph("Leandro Andreoli Santos Silva")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    assinaturaDiretor.editAsText()
      .setFontSize("14")
      .setFontFamily("Oswald")
    body.appendParagraph("Diretor\nCNPJ Nº 23.361.953/0001-80")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize("16")
      .setBold(true);
    doc.saveAndClose();
    var pdfBlob = DriveApp.getFileById(copia.getId()).getAs("application/pdf");
    pdfBlob.setName(copia.getName() + ".pdf");
    var pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    certificado.Certificado = "https://drive.google.com/uc?export=view&id=" + pdfFile.getId();
    DriveApp.getFileById(copia.getId()).setTrashed(true);
    certificado.Qrcode = qrUrl;
    certificado.Status = "GERADO";
    certificado.DtEmissao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    certificado = JSON.stringify(certificado);
    var retorno = SalvarCertificados(certificado);
    return retorno
  } catch (err) {
    Logger.log(err.message);
    return "Error: " + err.message;
  }
}
function SalvarCertificados(certificado) {
  try {
    var sheet = SpreadsheetApp.getActive();
    var GerarCertificadoSheet = sheet.getSheetByName("GERARCERTIFICADO");
    var ultimaLinha = GerarCertificadoSheet.getLastRow();
    var ultimaColuna = GerarCertificadoSheet.getLastColumn();
    certificado = JSON.parse(certificado);
    if (certificado['IdCertificado'] == '') {
      certificado['IdCertificado'] = BuscarMaiorId("GERARCERTIFICADO");
      certificado["Status"] = "PENDENTE"
      var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
      var dados = IndiceCabecalho(cabecalho, certificado);
      GerarCertificadoSheet.getRange(ultimaLinha + 1, 1, 1, dados.length).setValues([dados]);
      return "Certificado salvo com sucesso!"
    } else {
      for (var i = 2; i <= ultimaLinha; i++) {
        if (certificado['IdCertificado'] == GerarCertificadoSheet.getRange(i, 1).getValue()) {
          var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
          var dados = IndiceCabecalho(cabecalho, certificado);
          GerarCertificadoSheet.getRange(i, 1, 1, dados.length).setValues([dados]);
          return "Certificado atualizado com sucesso!"
        }
      }
    }
  } catch (err) {
    Logger.log(err);
    return "Error: " + err.message;
  }
}
function AtualizarCertificado(certificado) {
  try {
    certificado = JSON.parse(certificado);
    certificado["Status"] = "PENDENTE";
    certificado["Codigo"] = "";
    certificado["Envio"] = "";
    certificado["Qrcode"] = "";
    certificado["ChCurso"] = "";
    certificado["DtEmissao"] = "";
    certificado["Certificado"] = "";
    SalvarCertificados(certificado);
  } catch (e) {
    Logger.log(e.message);
    return e.message
  }
}
function BuscarCertificados() {
  var retorno = [];
  var sheet = SpreadsheetApp.getActive();
  var GerarCertificadoSheet = sheet.getSheetByName("GERARCERTIFICADO");
  var ultimaLinha = GerarCertificadoSheet.getLastRow();
  var ultimaColuna = GerarCertificadoSheet.getLastColumn();
  var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  var dados = GerarCertificadoSheet.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
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

function GerarCertificados() {
  try {
    var count = 0;
    var certificados = BuscarCertificados();
    certificados.forEach((certificado) => {
      if (certificado.Status == "PENDENTE") {
        GerarCertificado(certificado.IdCertificado);
        count++
      }
    })
    if (count == 0) {
      return `Foi gerado nenhum certificado. Não existem pendências!`
    } else if (count == 1) {
      return "Foi gerado 1 certificado!"
    } else {
      return `Foi gerado ${count} certificados!`
    }
  } catch (err) {
    console.log(err.message)
  }
}
function BuscarCertificado(idCertificado) {
  try {
    var sheet = SpreadsheetApp.getActive();
    var GerarCertificadoSheet = sheet.getSheetByName("GERARCERTIFICADO");
    var ultimaLinha = GerarCertificadoSheet.getLastRow();
    var ultimaColuna = GerarCertificadoSheet.getLastColumn();
    for (var i = 2; i <= ultimaLinha; i++) {
      if (GerarCertificadoSheet.getRange(i, 1).getValue() == idCertificado) {
        var retorno = {}
        var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var dados = GerarCertificadoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
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
function DeletarCertificado(idCertificado) {
  var sheet = SpreadsheetApp.getActive();
  var GerarCertificadoSheet = sheet.getSheetByName("GERARCERTIFICADO");
  var ultimaLinha = GerarCertificadoSheet.getLastRow();
  for (var i = 2; i <= ultimaLinha; i++) {
    if (GerarCertificadoSheet.getRange(i, 1).getValue() == idCertificado) {
      GerarCertificadoSheet.deleteRow(i);
      return "Certificado deletado com sucesso!";
    }
  }
  return "Certificado não encontrado!";
}

function EnviarCertificadoPorEmail(idCertificado) {
  try {
    var certificado = BuscarCertificado(idCertificado);
    certificado = JSON.parse(certificado);
    var aluno = BuscarAluno(certificado.IdAluno)
    aluno = JSON.parse(aluno);
    var fileId = certificado.Certificado.match(/id=([a-zA-Z0-9_-]+)/)[1];
    var pdfFile = DriveApp.getFileById(fileId);
    var destinatario = certificado.Email;
    var assunto = "Seu Certificado - " + certificado.Curso;
    var corpo = `
      Olá ${aluno.Genero == "Masculino" ? "Sr." : "Sra."} ${certificado.Nome},

      Parabéns pela conclusão do curso "${certificado.Curso}".
      Segue em anexo o seu certificado.
      
      Código do certificado: ${certificado.Codigo}
      Data de emissão: ${new Date(certificado.DtEmissao).toLocaleDateString()}

      Atenciosamente,
      Lar, Treinamentos e Projetos Elétricos. 
    `;
    GmailApp.sendEmail(destinatario, assunto, corpo, {
      attachments: [pdfFile.getAs("application/pdf")]
    })
    return "Email enviado com sucesso para " + destinatario;
  } catch (err) {
    return "Erro: " + err.message;
  }
}
function BuscarCertificadoPorCodigo(codigo) {
  var GerarCertificadoSheet = SpreadsheetApp.getActive().getSheetByName("GERARCERTIFICADO");
  var ultimaLinha = GerarCertificadoSheet.getLastRow();
  var ultimaColuna = GerarCertificadoSheet.getLastColumn();
  var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  for (var i = 2; i <= ultimaLinha; i++) {
    if (GerarCertificadoSheet.getRange(i, cabecalho.indexOf("Codigo") + 1).getValue() == codigo) {
      var retorno = {}
      var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
      var dados = GerarCertificadoSheet.getRange(i, 1, 1, ultimaColuna).getValues()[0];
      cabecalho.forEach((item, index) => {
        retorno[item] = dados[index]
      })
      retorno = JSON.stringify(retorno);
      return retorno;
    }
  }
  return null
}

function DeletarCertificadoPorCurso(IdCurso) {
  let count = 0;
  var GerarCertificadoSheet = SpreadsheetApp.getActive().getSheetByName("GERARCERTIFICADO");
  var ultimaLinha = GerarCertificadoSheet.getLastRow();
  var ultimaColuna = GerarCertificadoSheet.getLastColumn();
  var cabecalho = GerarCertificadoSheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  for (var i = ultimaLinha; i >= 2; i--) {
    if (GerarCertificadoSheet.getRange(i, cabecalho.indexOf("IdCursoTurma") + 1).getValue() == IdCurso) {
      GerarCertificadoSheet.deleteRow(i);
      count++
    }
  }
  if (count > 0) {
    return `Foram deletados ${count} registros de conteúdos programados para esse curso!`
  } else {
    return 'Nenhum conteúdo programático encontrado!'
  }
}