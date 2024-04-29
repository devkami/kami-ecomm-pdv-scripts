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
  };
  
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
}

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
  };
  
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
}


function lerColuna() {
  // Obter a planilha ativa
  var folha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Preços do Tiny");
  
  // Especificar a coluna que você deseja ler (por exemplo, coluna A)
  var coluna = folha.getRange('A2:A');
  
  // Obter os valores na coluna
  var valores = coluna.getValues();
  
  // Iterar pelos valores e obter informações de preço, preço de custo, marca, descrição e custo
  for (var i = 0; i < valores.length; i++) {
    var sku = valores[i][0];    
    Logger.log("Lendo SKU:" + sku)
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




function obterCustoProduto(sku) {
  var url = 'https://api.tiny.com.br/api2/produtos.pesquisa.php';
  var token = 'TINY_API_TOKEN'; // Substitua pelo seu token real 
  
  var params = {
    method: 'post',
    payload: {
      token: token,
      pesquisa: sku,
      formato: 'JSON'
    }
  };
  
  var response = UrlFetchApp.fetch(url, params);
  
  if (response.getResponseCode() === 200) {
    var jsonResponse = response.getContentText();
    var jsonData = JSON.parse(jsonResponse);
    // Verifica se a resposta contém os dados do produto
    if (jsonData.retorno && jsonData.retorno.produtos && jsonData.retorno.produtos.length > 0) {
      var produto = jsonData.retorno.produtos[0].produto;
      // Retorna o preço de custo e o ID do produto
      return {
        preco_custo: produto.preco_custo, // Supõe-se que essa chave está correta
        id: produto.id
      };
    } else {
      Logger.log('Produto não encontrado para SKU: ' + sku);
      return null;
    }
  } else {
    Logger.log('Erro na solicitação: ' + response.getResponseCode());
    return null;
  }
}


function obterEstoqueProduto(id) {
  var url = 'https://api.tiny.com.br/api2/produto.obter.estoque.php';
  var token = 'TINY_API_TOKEN'; // Substitua pelo seu token real 
  var formato = 'JSON';

  // Construir o payload da solicitação
  var data = {
    'token': token,
    'id': id,
    'formato': formato
  };

  // Definir os parâmetros da solicitação HTTP, incluindo o método e o conteúdo
  var params = {
    'method': 'post',
    'payload': data
  };

  // Enviar a solicitação e capturar a resposta
  try {
    var response = UrlFetchApp.fetch(url, params);
    var content = response.getContentText();

    // Verificar se a resposta está OK
    if (response.getResponseCode() === 200) {
      // Se sim, converter o texto da resposta para JSON
      var jsonResponse = JSON.parse(content);
      return jsonResponse;
    } else {
      // Se não, logar o código de resposta HTTP e retornar nulo
      console.error('Erro na solicitação: Código ' + response.getResponseCode());
      return null;
    }
  } catch (e) {
    // Capturar e logar exceções
    console.error('Exceção ao fazer solicitação: ' + e.toString());
    return null;
  }
}