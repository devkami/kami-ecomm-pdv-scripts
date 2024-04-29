/**
 * Nome do Script: Gerenciador de Atualização de Produtos para Google Sheets
 * Descrição:
 * Este script automatiza o processo de atualização de informações de produtos na planilha "Preços do Tiny" do Google Sheets.
 * Ele gerencia as interações com a API para buscar os detalhes mais recentes dos produtos e atualiza a planilha de acordo.
 * O script utiliza programação orientada a objetos para gerenciar os dados e atualizações dos produtos de forma eficiente. O tratamento de erros
 * está integrado para gerenciar questões relacionadas à conectividade da API ou discrepâncias de dados.
 *
 * Funções:
 * - Produto(sku, marca, nome, peso_liquido, preco_custo, preco): Construtor para criar um objeto de produto.
 * - GerenciadorProduto(nomePlanilha): Construtor que inicializa um gerenciador de produtos para uma planilha do Google Sheets específica.
 * - atualizarProdutosPlanilha(): Função principal para processar e atualizar todas as entradas de produtos na planilha.
 * - carregarProdutoDaPlanilha(linha): Carrega dados de produtos de uma linha específica na planilha.
 * - buscarDadosProduto(sku): Busca dados atualizados do produto de uma API externa usando o SKU do produto.
 * - atualizarPlanilhaSeNecessario(linha, diferencas): Atualiza a planilha se houver diferenças entre os dados existentes e os novos.
 * - inicializarProgresso(), atualizarProgresso() e atualizarDataHoraFinal(): Gerenciam a exibição do progresso da execução e do tempo dentro da planilha.
 *
 * Uso:
 * - O script é vinculado a um arquivo do Google Sheets e pode ser executado diretamente do editor do Google Apps Script.
 * - Requer configuração adequada de tokens OAuth e permissões para acessar tanto o Google Sheets quanto a API de produtos externa.
 *
 * Exemplo:
 * Para usar este script, basta configurar gatilhos no ambiente do Google Apps Script para executar a função atualizarProdutosPlanilha
 * conforme necessário (por exemplo, gatilhos baseados em tempo para atualizações diárias).
 *
 * Tratamento de Erros:
 * - O script inclui tratamento de erros para reg/istrar problemas relacionados a problemas de rede, falhas da API ou dados inválidos.
 *   Os erros são registrados na facilidade de logs do Google Apps Script e podem ser revisados para solução de problemas.
 *
 * Autor: Maicon de Menezes <maicondmenezes@gmail.com>
 * Data de Criação: 27/04/2024
 * Versão: 0.1.0
 */

/**
 * Classe Produto
 * Representa um produto com atributos específicos extraídos ou destinados a uma planilha do Google Sheets.
 * 
 * @param {string} sku - SKU do produto.
 * @param {string} marca - Marca do produto.
 * @param {string} nome - Descrição do produto.
 * @param {number} peso_liquido - Peso líquido do produto em quilogramas.
 * @param {number} preco_custo - Preço de custo do produto.
 * @param {number} preco - Preço sugerido de venda do produto.
 */
class Produto {
  constructor(sku, marca, nome, peso_liquido, preco_custo, preco) {
    this.sku = sku;
    this.marca = marca;
    this.nome = nome;
    this.peso_liquido = peso_liquido;
    this.preco_custo = preco_custo;
    this.preco = preco;
  }

  /**
   * Compara este produto com outro para identificar diferenças nos atributos.
   * 
   * @param {Produto} outroProduto - O produto com o qual comparar.
   * @returns {Object} Objeto contendo apenas os atributos que diferem com os valores do outroProduto.
   */
  comparar(outroProduto) {
    let diferencas = {};
    Object.keys(this).forEach(chave => {
      if (this[chave] !== outroProduto[chave]) {
        diferencas[chave] = outroProduto[chave];
      }
    });
    return diferencas;
  }
}

/**
 * Classe GerenciadorProduto
 * Gerencia a atualização de produtos em uma planilha do Google Sheets.
 * 
 * @param {string} nomePlanilha - Nome da planilha dentro do Google Spreadsheet.
 */
class GerenciadorProduto {
  constructor(nomePlanilha) {
    this.planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha);
    this.linhaAtual = 0;
    this.totalLinhas = 0;
    this.dataHoraInicio = new Date();
    this.dataHoraFim = null;
    this.apiToken = 'TINY_API_TOKEN'; // Substitua pelo seu token real 
    this.baseUrl = 'https://api.tiny.com.br/api2'; 
    this.inicializarProgresso();
  }

  /**
   * Pesquisa um produto pelo SKU usando a API externa.
   * Realiza uma requisição POST para obter dados do produto com base no seu SKU.
   * 
   * @param {string} sku - O SKU do produto a ser pesquisado.
   * @returns {Object|null} Retorna os dados do produto como um objeto JSON se bem-sucedido, ou null se houver falha.
   * @throws {Error} Lança um erro se a requisição não retornar status 200.
   */
  pesquisarProduto(sku) {
    var url = `${this.baseUrl}/produtos.pesquisa.php`; 
    var params = {
      method: 'post',
      payload: {
        token: this.apiToken, 
        pesquisa: sku,
        formato: 'JSON'
      }
    };

    try {
      var response = UrlFetchApp.fetch(url, params);
      if (response.getResponseCode() === 200) {
        var jsonData = JSON.parse(response.getContentText());
        return jsonData;
      } else {
        throw new Error('Response returned with status: ' + response.getResponseCode());
      }
    } catch (error) {
      Logger.log('Erro na pesquisa de produto: ' + error.toString());
      return null;
    }
  }

  /**
   * Obtém detalhes de um produto específico pelo ID usando a API externa.
   * Realiza uma requisição POST para obter detalhes completos do produto.
   * 
   * @param {string} id - O ID do produto para obter detalhes.
   * @returns {Object|null} Retorna os dados detalhados do produto como um objeto JSON se bem-sucedido, ou null se houver falha.
   * @throws {Error} Lança um erro se a requisição não retornar status 200.
   */
  obterProduto(id) {
    var url = `${this.baseUrl}/produto.obter.php`;
    var params = {
      method: 'post',
      payload: {
        token: this.apiToken,
        id: id,
        formato: 'JSON'
      }
    };

    try {
      var response = UrlFetchApp.fetch(url, params);
      if (response.getResponseCode() === 200) {
        var jsonData = JSON.parse(response.getContentText());
        return jsonData;
      } else {
        throw new Error('Response returned with status: ' + response.getResponseCode());
      }
    } catch (error) {
      Logger.log('Erro ao obter produto: ' + error.toString());
      return null;
    }
  }

  /**
   * Atualiza todos os produtos na planilha, manipulando a autenticação e a paginação de dados se necessário.
   */
  atualizarProdutosPlanilha() {
    const coluna = this.planilha.getRange('A2:A' + this.planilha.getLastRow());
    const skus = coluna.getValues();
    this.totalLinhas = skus.length;
  
    skus.forEach((valor, indice) => {
      this.linhaAtual = indice + 2;
      this.atualizarProgresso();
      const valorSKU = valor[0];
      if (valorSKU !== "" && valorSKU !== null) {
        let produtoAtual = this.carregarProdutoDaPlanilha(this.linhaAtual);
        let produtoPesquisado = this.pesquisarProduto(produtoAtual.sku);
        if (produtoPesquisado && produtoPesquisado.retorno && produtoPesquisado.retorno.produtos && produtoPesquisado.retorno.produtos.length > 0) {
          let detalhesProduto = this.obterProduto(produtoPesquisado.retorno.produtos[0].produto.id);
          if (detalhesProduto && detalhesProduto.retorno) {
            let novoProduto = new Produto(
              detalhesProduto.retorno.produto.sku,
              detalhesProduto.retorno.produto.marca,
              detalhesProduto.retorno.produto.nome,
              detalhesProduto.retorno.produto.peso_liquido,
              detalhesProduto.retorno.produto.preco_custo,
              detalhesProduto.retorno.produto.preco
            );
            let diferencas = produtoAtual.comparar(novoProduto);
            this.atualizarPlanilhaSeNecessario(this.linhaAtual, diferencas);
          } else {
            this.atualizarMetadados(this.linhaAtual, "Erro: Falha ao obter detalhes do produto.", []);
          }
        } else {
          this.atualizarMetadados(this.linhaAtual, "Erro: Falha ao buscar produto.", []);
        }
      }
    });
  
    this.dataHoraFim = new Date();
    this.atualizarDataHoraFinal();
  }


  /**
   * Carrega um produto da planilha a partir de uma linha especificada.
   * 
   * @param {number} linha - A linha da planilha de onde carregar o produto.
   * @returns {Produto} Uma instância da classe Produto com dados carregados da planilha.
   */
  carregarProdutoDaPlanilha(linha) {
    const valores = this.planilha.getRange('A' + linha + ':F' + linha).getValues()[0];
    return new Produto(valores[0], valores[1], valores[2], valores[3], valores[4], valores[5]);
  }

  /**
   * Busca dados atualizados de um produto usando a API externa.
   * 
   * @param {string} sku - O SKU do produto a ser buscado.
   * @returns {Produto|null} Uma nova instância de Produto com dados atualizados ou null se houver falha.
   */
  buscarDadosProduto(sku) {
    try {
      const infoProduto = this.pesquisarProduto(sku);
      if (infoProduto && infoProduto.retorno.produtos.length > 0) {
        const dadosProduto = infoProduto.retorno.produtos[0].produto;
        return new Produto(sku, dadosProduto.marca, dadosProduto.nome, dadosProduto.peso_liquido, dadosProduto.preco_custo, dadosProduto.preco);
      }
    } catch (erro) {
      Logger.log(`Erro ao buscar dados para o SKU: ${sku} com erro: ${erro.toString()}`);
      return null;
    }
  }

  /**
   * Atualiza a planilha se houver diferenças significativas entre os dados existentes e os novos.
   *
   * @param {number} linha - A linha da planilha a ser atualizada.
   * @param {Object} diferencas - Um objeto contendo as diferenças detectadas.
  */
  atualizarPlanilhaSeNecessario(linha, diferencas) {
      // Ensure the keys correspond to the correct columns: B - marca, C - nome, D - peso_liquido, E - preco_custo, F - preco
      if (diferencas.marca) {
          this.planilha.getRange('B' + linha).setValue(diferencas.marca);
      }
      if (diferencas.nome) {
          this.planilha.getRange('C' + linha).setValue(diferencas.nome);
      }
      if (diferencas.peso_liquido) {
          this.planilha.getRange('D' + linha).setValue(diferencas.peso_liquido);
      }
      if (diferencas.preco_custo) {
          this.planilha.getRange('E' + linha).setValue(diferencas.preco_custo);
      }
      if (diferencas.preco) {
          this.planilha.getRange('F' + linha).setValue(diferencas.preco);
      }
  
      // Debugging: Log differences to verify what is being updated
      Logger.log("Updating line: " + linha + " with differences: " + JSON.stringify(diferencas));
  
      const camposAtualizados = Object.keys(diferencas);
      const mensagemResultado = camposAtualizados.length > 0 ? `Campos atualizados com sucesso: [${camposAtualizados.join(', ')}]` : "Sem alterações";
      this.atualizarMetadados(linha, mensagemResultado, camposAtualizados);
  }


  /**
   * Inicializa a exibição do progresso da execução na planilha.
   */
  inicializarProgresso() {
    const formatoDataHora = Utilities.formatDate(this.dataHoraInicio, "GMT-03:00", "dd/MM/yyyy HH:mm:ss");
    this.planilha.getRange('L2').setValue(formatoDataHora);
  }

  /**
   * Atualiza o progresso da execução na planilha com a linha atual e o total de linhas.
   */
  atualizarProgresso() {
    this.planilha.getRange('K5').setValue(this.linhaAtual - 1);
    this.planilha.getRange('M5').setValue(this.totalLinhas);
  }

  /**
   * Registra a data e hora de término da execução na planilha.
   */
  atualizarDataHoraFinal() {
    const formatoDataHora = Utilities.formatDate(this.dataHoraFim, "GMT-03:00", "dd/MM/yyyy HH:mm:ss");
    this.planilha.getRange('L3').setValue(formatoDataHora);
  }

  /**
   * Atualiza as informações de metadados na planilha com a data, hora e resultado da atualização.
   * Inclui detalhes sobre quais campos foram atualizados.
   *
   * @param {number} linha - A linha onde os metadados serão atualizados.
   * @param {string} mensagemResultado - Mensagem indicando o resultado da atualização.
   * @param {Array} camposAtualizados - Lista dos campos que foram atualizados.
   */
  atualizarMetadados(linha, mensagemResultado, camposAtualizados) {
    const agora = new Date();
    const formatoData = Utilities.formatDate(agora, "GMT-03:00", "dd/MM/yyyy");
    const formatoHora = Utilities.formatDate(agora, "GMT-03:00", "HH:mm:ss");
  
    // Incluir detalhes sobre os campos atualizados na mensagem de resultado
    let detalheCampos = camposAtualizados.length > 0 ? " Campos atualizados: " + camposAtualizados.join(', ') : "";
    let mensagemCompleta = mensagemResultado + detalheCampos;
  
    const intervaloMetadados = this.planilha.getRange('G' + linha + ':I' + linha);
    intervaloMetadados.setValues([[formatoData, formatoHora, mensagemCompleta]]);
  }

  /**
   * Verifica a necessidade de atualização de cada linha na planilha com base na coluna 'G' de última atualização.
   * Atualiza a linha se a coluna 'G' estiver vazia ou se a data for menor que o dia atual.
   * 
   * @returns {void} Não retorna nada.
   * @throws {Error} Lança um erro se a operação de atualização falhar.
   */
  atualizarProdutosAutomaticamente() {
    const hoje = Utilities.formatDate(new Date(), "GMT-03:00", "yyyy-MM-dd");
    const ultimaLinha = this.planilha.getLastRow();
    const dadosSKU = this.planilha.getRange('A2:A' + ultimaLinha).getValues();
    const dadosUltimaAtualizacao = this.planilha.getRange('G2:G' + ultimaLinha).getValues();
  
    this.totalLinhas = dadosSKU.length;
  
    dadosSKU.forEach((row, index) => {
        this.linhaAtual = index + 2;
        const valorSKU = row[0];
        const dataUltimaAtualizacao = dadosUltimaAtualizacao[index][0];
        if (!dataUltimaAtualizacao || dataUltimaAtualizacao < hoje) {
            this.atualizarProdutoSeNecessario(valorSKU);
        }
    });
  
    this.dataHoraFim = new Date();
    this.atualizarDataHoraFinal();
  }

  /**
   * Atualiza os detalhes de um produto específico na planilha se necessário.
   * Chama métodos para pesquisar o produto na API externa e atualiza a planilha com os novos dados se houver diferenças.
   * 
   * @param {string} sku - O SKU do produto a ser atualizado.
   * @returns {void} Não retorna nada.
   * @throws {Error} Lança um erro se não conseguir atualizar os dados do produto na planilha.
   */
  atualizarProdutoSeNecessario(sku) {
    let produtoAtual = this.carregarProdutoDaPlanilha(this.linhaAtual);
    let produtoPesquisado = this.pesquisarProduto(sku);
    if (produtoPesquisado && produtoPesquisado.retorno && produtoPesquisado.retorno.produtos && produtoPesquisado.retorno.produtos.length > 0) {
      let detalhesProduto = this.obterProduto(produtoPesquisado.retorno.produtos[0].produto.id);
      if (detalhesProduto && detalhesProduto.retorno) {
        let novoProduto = new Produto(
          detalhesProduto.retorno.produto.sku,
          detalhesProduto.retorno.produto.marca,
          detalhesProduto.retorno.produto.nome,
          detalhesProduto.retorno.produto.peso_liquido,
          detalhesProduto.retorno.produto.preco_custo,
          detalhesProduto.retorno.produto.preco
        );
        let diferencas = produtoAtual.comparar(novoProduto);
        this.atualizarPlanilhaSeNecessario(this.linhaAtual, diferencas);
      } else {
        this.atualizarMetadados(this.linhaAtual, "Erro: Falha ao obter detalhes do produto.", []);
      }
    } else {
      this.atualizarMetadados(this.linhaAtual, "Erro: Falha ao buscar produto.", []);
    }
  }

}

/**
 * Atualiza os produtos na planilha especificada.
 * Esta função pode ser acionada por um objeto de desenho que atua como botão em diferentes planilhas.
 * Cada chamada da função pode especificar o nome da planilha a ser processada.
 *
 * @param {string} nomePlanilha - Nome da planilha a ser atualizada.
 */
function atualizarProdutos(nomePlanilha) {
  const gerenciador = new GerenciadorProduto(nomePlanilha);
  gerenciador.atualizarProdutosPlanilha();
}

/**
 * Função agendada para executar a atualização automática dos produtos na planilha diariamente à meia-noite.
 * Essa função é disparada por um gatilho programado.
 * 
 * @returns {void} Não retorna nada.
 * @throws {Error} Lança um erro se a atualização programada falhar.
 */
function autoAtualizarProdutos_Precos_do_Tiny() {
  const gerenciador = new GerenciadorProduto("Preços do Tiny");
  gerenciador.atualizarProdutosAutomaticamente();
}

/**
 * Função agendada para executar a atualização automática dos produtos na planilha diariamente à meia-noite.
 * Essa função é disparada por um gatilho programado.
 * 
 * @returns {void} Não retorna nada.
 * @throws {Error} Lança um erro se a atualização programada falhar.
 */
function autoAtualizarProdutos_test_DEV() {
  const gerenciador = new GerenciadorProduto("teste DEV");
  gerenciador.atualizarProdutosAutomaticamente();
}



function atualizarProdutos_teste_DEV(){
  atualizarProdutos("teste DEV")
}

function atualizarProdutos_Precos_do_Tiny(){
  atualizarProdutos("Preços do Tiny")
}

