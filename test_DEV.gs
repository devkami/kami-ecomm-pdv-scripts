function pesquisarProduto(sku) {
  var url = 'https://api.tiny.com.br/api2/produtos.pesquisa.php';
  var token = 'TINY_API_TOKEN'; // Substitua pelo seu token real
  
  // Construir os parâmetros da solicitação
  var params = {
    method: 'post',
    payload: {
      token: token,
      pesquisa: sku,
      formato: 'JSON'
    }
  }

  // Realizar a solicitação HTTP
  var response = UrlFetchApp.fetch(url, params);
  
  // Verificar se a solicitação foi bem-sucedida
  if (response.getResponseCode() === 200) {
    var jsonResponse = response.getContentText();
    var jsonData = JSON.parse(jsonResponse);
    
    // Exibir a resposta (opcional)
    // Logger.log(jsonData);
    
    return jsonData;
  } else {
    // Tratar erros, se necessário
    Logger.log('Erro na solicitação: ' + response.getResponseCode());
    return null;
  }
};

function obterProduto(id) {
  var url = 'https://api.tiny.com.br/api2/produto.obter.php';
  var token = 'TINY_API_TOKEN'; // Substitua pelo seu token real
  
  // Construir os parâmetros da solicitação
  var params = {
    method: 'post',
    payload: {
      token: token,
      id: id,
      formato: 'JSON'
    }
  }
};


var teste2 = pesquisarProduto('BR0001');
var teste1 = obterProduto('743004367');
console.log(teste1);




function teste() {
  // Obter a planilha ativa
  var folha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("teste DEV");

  // Especificar a coluna que você deseja ler (por exemplo, coluna A)
  var coluna = folha.getRange('A2:A');

  // Obter os valores na coluna
  var valores = coluna.getValues();

  // Iterar pelos valores e obter informações de preço, preço de custo, marca, descrição e custo
  for (var i = 0; i < valores.length; i++) {
    var sku = valores[i][0];

    // Verificar se a célula contém algum valor
    if (sku !== "" && sku !== null) {
      // Verificar se as colunas de marca, descrição, peso e custo estão vazias
      var linha = i + 2;  // Linha atual na folha (começando da linha 2)
      var marcaCelula = folha.getRange('B' + linha).getValue();
      var descricaoCelula = folha.getRange('C' + linha).getValue();
      var pesoCelula = folha.getRange('D' + linha).getValue();
      var custoCelula = folha.getRange('E' + linha).getValue();

      if (marcaCelula === "" && descricaoCelula === "" && pesoCelula === "" && custoCelula === "") {
        // Todas as colunas estão vazias, fazer uma nova requisição
        var produtoInfo = pesquisarProduto(sku);

        // Se o produtoInfo for nulo, significa que o produto não foi encontrado
        if (produtoInfo && produtoInfo.retorno.produtos && produtoInfo.retorno.produtos.length > 0) {
          // Obter informações adicionais usando o ID do produto
          var idProduto = produtoInfo.retorno.produtos[0].produto.id; // Ajuste conforme necessário

          // Verificar se idProduto é definido antes de usar
          if (idProduto !== undefined && idProduto !== null) {
            var infoAdicional = obterProduto(idProduto);

            if (infoAdicional) {
              // Escrever os valores nas colunas B, C, D e E
              folha.getRange('B' + linha).setValue(infoAdicional.retorno.produto.marca);
              folha.getRange('C' + linha).setValue(infoAdicional.retorno.produto.nome);
              folha.getRange('D' + linha).setValue(infoAdicional.retorno.produto.peso_liquido);
              folha.getRange('E' + linha).setValue(infoAdicional.retorno.produto.preco_custo);

              Logger.log("SKU: " + sku + ", Marca: " + infoAdicional.retorno.produto.marca + ", Descrição: " + infoAdicional.retorno.produto.nome + ", Peso: " + infoAdicional.retorno.produto.peso_liquido + ", Custo: " + infoAdicional.retorno.produto.preco_custo);
            } else {
              Logger.log("Não foi possível obter informações adicionais para SKU: " + sku);
              // Adicionar mensagem de erro na coluna de marca
              folha.getRange('B' + linha).setValue("A consulta não retornou registros");
            }
          } else {
            Logger.log("ID do produto não está definido para SKU: " + sku);
          }
        } else {
          Logger.log("Não foi possível obter informações para SKU: " + sku);
          // Adicionar mensagem de erro na coluna de marca
          folha.getRange('B' + linha).setValue("A consulta não retornou registros");
        }
      } else {
        Logger.log("Alguma informação já existe para SKU: " + sku);
      }
    }
  }
}

function lerColuna() {
  // Obter a planilha ativa
  var folha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("teste DEV");

  // Especificar a coluna que você deseja ler (por exemplo, coluna A)
  var coluna = folha.getRange('A2:A');

  // Obter os valores na coluna
  var valores = coluna.getValues();

  // Iterar pelos valores e obter informações de preço, preço de custo, marca, descrição e custo
  for (var i = 0; i < valores.length; i++) {
    var sku = valores[i][0];

    // Verificar se a célula contém algum valor
    if (sku !== "" && sku !== null) {
      // Verificar se as colunas de marca, descrição, peso e custo estão vazias
      var linha = i + 2;  // Linha atual na folha (começando da linha 2)
      var marcaCelula = folha.getRange('B' + linha).getValue();
      var descricaoCelula = folha.getRange('C' + linha).getValue();
      var pesoCelula = folha.getRange('D' + linha).getValue();
      var custoCelula = folha.getRange('E' + linha).getValue();

      if (marcaCelula === "" && descricaoCelula === "" && pesoCelula === "" && custoCelula === "") {
        // Todas as colunas estão vazias, fazer uma nova requisição
        var produtoInfo = pesquisarProduto(sku);
        


        // Se o produtoInfo for nulo, significa que o produto não foi encontrado
        if (produtoInfo && produtoInfo.retorno.produtos && produtoInfo.retorno.produtos.length > 0) {
          // Obter informações adicionais usando o ID do produto
          var idProduto = produtoInfo.retorno.produtos[0].produto.id; // Ajuste conforme necessário

          // Verificar se idProduto é definido antes de usar
          if (idProduto !== undefined && idProduto !== null) {
            var infoAdicional = obterProduto(idProduto);
            if (infoAdicional) {
              // Escrever os valores nas colunas B, C, D e E
              folha.getRange('B' + linha).setValue(infoAdicional.retorno.produto.marca);
              folha.getRange('C' + linha).setValue(infoAdicional.retorno.produto.nome);
              folha.getRange('D' + linha).setValue(infoAdicional.retorno.produto.peso_liquido);
              folha.getRange('E' + linha).setValue(infoAdicional.retorno.produto.preco_custo);
              folha.getRange('F' + linha).setValue(infoAdicional.retorno.produto.preco);

              Logger.log("SKU: " + sku + ", Marca: " + infoAdicional.retorno.produto.marca + ", Descrição: " + infoAdicional.retorno.produto.nome + ", Peso: " + infoAdicional.retorno.produto.peso_liquido + ", Custo: " + infoAdicional.retorno.produto.preco_custo);
            } else {
              Logger.log("Não foi possível obter informações adicionais para SKU: " + sku);
              // Adicionar mensagem de erro na coluna de marca
              folha.getRange('B' + linha).setValue("A consulta não retornou registros");
            }
          } else {
            Logger.log("ID do produto não está definido para SKU: " + sku);
          }
        } else {
          Logger.log("Não foi possível obter informações para SKU: " + sku);
          // Adicionar mensagem de erro na coluna de marca
          folha.getRange('B' + linha).setValue("A consulta não retornou registros");
        }
      } else {
        Logger.log("Alguma informação já existe para SKU: " + sku);
      }
    }
  }
}

lerColuna();

