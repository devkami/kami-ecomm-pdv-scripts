function atualizarSaldoEstoque() {
  var folha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque Tiny");
  var colunaSKUs = folha.getRange('A2:A' + folha.getLastRow()).getValues();
  
  for (var i = 0; i < colunaSKUs.length; i++) {
    var sku = colunaSKUs[i][0];
    if (sku !== "" && sku !== null) {
      var linha = i + 2;

      // Obter o ID do produto com base no SKU
      var custoProduto = obterCustoProduto(sku); 
      if (custoProduto && custoProduto.id) {
        // Obter informações de estoque usando o ID do produto
        var estoqueProduto = obterEstoqueProduto(custoProduto.id);
        if (estoqueProduto && estoqueProduto.retorno.status === 'OK' && estoqueProduto.retorno.produto) {
          var produto = estoqueProduto.retorno.produto;
          var depositos = produto.depositos;

          // Atualizar o saldo de estoque para cada depósito
          depositos.forEach(function(deposito, index) {
            var colunaSaldo = String.fromCharCode(67 + index);
            var saldo = deposito.deposito.saldo;
            var nomeDeposito = deposito.deposito.nome;
            var celulaSaldoAtual = folha.getRange(colunaSaldo + linha).getValue();

            // Verifica se o saldo na planilha difere do saldo retornado pela API
            if (celulaSaldoAtual !== saldo) {
              folha.getRange(colunaSaldo + linha).setValue(saldo);
              Logger.log("Saldo atualizado para SKU: " + sku + ", Depósito: " + nomeDeposito + ", Novo Saldo: " + saldo);
            }
          });
        } else {
          Logger.log("Não foi possível obter informações de estoque para SKU: " + sku);
          // Pode-se considerar adicionar alguma indicação de erro na planilha
        }
      } else {
        Logger.log("Não foi possível obter o ID para SKU: " + sku);
        // Pode-se considerar adicionar alguma indicação de erro na planilha
      }
    }
  }
}