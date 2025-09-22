function FormAluno() {
  var form = HtmlService.createTemplateFromFile("FormAluno");
  var showForm = form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  showForm.setTitle("Cadastro de Alunos").setHeight(300).setWidth(500);
}

function FormCadastroAluno(IdAluno) {
  if (IdAluno == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("AlunoCadastroForm");
    form.setWidth(800)
    form.setHeight(600)
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Cadastro de Aluno");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("AlunoCadastroForm").append(`<div id="AlunoEdit" style='display:none' >${IdAluno}</div>`)
    form.setWidth(800)
    form.setHeight(600)
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Editar registro de Aluno");
  }

}
function FormListarAluno() {
  var form = HtmlService.createHtmlOutputFromFile("AlunoListaForm");
  form.setWidth(800)
  form.setHeight(600)
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Alunos");

}
function FormCadastroCursoTurma(IdCursoTurma) {
  if (IdCursoTurma == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("CursoTurmaCadastroForm");
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Cadastro do Curso");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("CursoTurmaCadastroForm").append(`<div id="CursoTurmaEdit" style='display:none' >${IdCursoTurma}</div>`);
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Editar registro do Curso");
  }
}
function FormListarCursoTurma() {
  var form = HtmlService.createHtmlOutputFromFile("CursoTurmaListaForm");
  form.setWidth(800);
  form.setHeight(600);
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Cursos turmas");
}
function FormCadastrarModeloCertificado(IdModeloCertificado) {
  if (IdModeloCertificado == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("ModelosCertificadosCadastroForm");
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Registrar Modelo");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("ModelosCertificadosCadastroForm").append(`<div id="ModeloCertificadoEdit" style='display:none' >${IdModeloCertificado}</div>`);
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Editar registro do Modelo");
  }
}
function FormListarModelosCertificados() {
  var form = HtmlService.createHtmlOutputFromFile("ModelosCertificadosListaForm");
  form.setWidth(800);
  form.setHeight(600);
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Modelos")
}
function FormCadastrarCertificado(IdCertificado) {
  if (IdCertificado == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("CertificadosCadastroForm");
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Registrar Certificado");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("CertificadosCadastroForm")
      .append(`<div id="CertificadoEdit" style='display:none' >${IdCertificado}</div>`);
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Editar registro do certificado");
  }
}
function FormListarCertificados() {
  var form = HtmlService.createHtmlOutputFromFile("CertificadosListaForm");
  form.setWidth(800);
  form.setHeight(600);
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Certificados")
}
function FormCadastrarConteudoProgramatico(IdConteudoProgramatico) {
  if (IdConteudoProgramatico == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("ConteudoProgramaticoCadastroForm");
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Adicionar Conteudo Programático");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("ConteudoProgramaticoCadastroForm")
      .append(`<div id="ConteudoProgramaticoEdit" style='display:none' >${IdConteudoProgramatico}</div>`);
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Edição Conteudo Programático");
  }
}
function FormListarConteudosProgramaticos() {
  var form = HtmlService.createHtmlOutputFromFile("ConteudoProgramaticoListaForm");
  form.setWidth(800);
  form.setHeight(600);
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Conteúdos dos Cursos")
}
function FormCadastrarPalestrante(IdPalestrante) {
  if (IdPalestrante == undefined) {
    var form = HtmlService.createHtmlOutputFromFile("PalestranteCadastroForm");
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Adicionar Palestrante");
  } else {
    var form = HtmlService.createHtmlOutputFromFile("PalestranteCadastroForm")
      .append(`<div id="PalestranteEdit" style='display:none' >${IdPalestrante}</div>`);
    form.setWidth(800);
    form.setHeight(600);
    var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Edição de registro do Palestrante");
  }
}
function FormListarPalestrantes(){
  var form = HtmlService.createHtmlOutputFromFile("PalestranteListaForm");
  form.setWidth(800);
  form.setHeight(600);
  var ui = SpreadsheetApp.getUi().showModelessDialog(form, "Listagem de Palestrantes")
}