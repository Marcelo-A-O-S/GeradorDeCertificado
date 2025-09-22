function onOpen(){
    let ui = SpreadsheetApp.getUi();
    ui.createMenu("🧩Painel de Controle")
    .addItem("Ver menu lateral", "verMenuLateral")
    .addToUi();
}
function verMenuLateral(){
    let menu = HtmlService.createHtmlOutputFromFile("MenuLaretal").setTitle("Painel de Controle");
    let ui = SpreadsheetApp.getUi();
    ui.showSidebar(menu)
}