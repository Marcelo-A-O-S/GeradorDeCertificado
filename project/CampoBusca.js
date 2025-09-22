function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("CampoBuscaForm.html");
/*   var codigo = e.parameter.codigo;
  if (!codigo) {
    return ContentService.createTextOutput(
      JSON.stringify({ erro: "Informe um código válido." })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var certificado = BuscarCertificadoPorCodigo(codigo);
  if (certificado == null) {
    return ContentService.createTextOutput(
      JSON.stringify({ erro: "Certificado não encontrado." })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(certificado)
    .setMimeType(ContentService.MimeType.JSON); */
}
