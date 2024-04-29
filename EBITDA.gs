function calcularEBITDA() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Calcular EBITDA'); // Substitua 'SUA_ABA' pelo nome da sua aba
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Loop através das linhas do array, começando na segunda linha para ignorar o cabeçalho
  for (var i = 1; i < values.length; i++) {
    var pdv = values[i][2];
    var custo = values[i][3];
    var comissao = values[i][4];
    var frete = values[i][5];
    var insumo = values[i][6];
    var adm = values[i][7];
    var reversa = values[i][8];

    // Certifique-se de que os valores sejam numéricos
    pdv = parseFloat(pdv) || 0;
    custo = parseFloat(custo) || 0;
    comissao = parseFloat(comissao) || 0;
    frete = parseFloat(frete) || 0;
    insumo = parseFloat(insumo) || 0;
    adm = parseFloat(adm) || 0;
    reversa = parseFloat(reversa) || 0;

    // Calcula EBITDA R$ e %
    var ebitdaReais = pdv - custo - comissao - frete - insumo - adm - reversa;
    var ebitdaPercentual = (ebitdaReais / pdv) * 100;

    // Se EBITDA % < 4, ajustar PDV até que seja >= 4%
    while (ebitdaPercentual < 4 && pdv < 1000) { // Adicionei um limite para PDV para evitar loop infinito
      pdv += 0.10; // Ajuste o PDV
      comissao = pdv * 0.15;
      adm = pdv * 0.05;
      reversa = pdv * 0.003;
      ebitdaReais = pdv - custo - comissao - frete - insumo - adm - reversa;
      ebitdaPercentual = (ebitdaReais / pdv) * 100;
    }

    // Definir os valores calculados de volta na planilha
    sheet.getRange(i + 1, 3).setValue(pdv.toFixed(2)); // PDV atualizado na coluna C
    sheet.getRange(i + 1, 9).setValue(ebitdaReais.toFixed(2)); // EBITDA R$ na coluna I
    sheet.getRange(i + 1, 10).setValue(ebitdaPercentual.toFixed(2)); // EBITDA % na coluna J
  }
}



