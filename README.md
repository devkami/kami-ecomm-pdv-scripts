# KAMI CO E-Commerce PDV Scripts

## Índice

- [Sobre](#sobre)
- [Uso](#uso)
- [Contribuição](CONTRIBUTING.md)
- [Autores](#autores)

## Sobre <a name = "sobre"></a>

Este projeto inclui uma série de scripts desenvolvidos para otimizar e automatizar a gestão de estoque e preços de produtos em planilhas do Google Sheets, utilizadas por uma empresa de e-commerce. Os scripts interagem com APIs externas para atualizar informações sobre produtos, gerenciar estoques e preços, e são projetados para serem facilmente integrados ao ambiente de Google Apps Script. O objetivo é facilitar a atualização e manutenção de dados críticos de venda e estoque, reduzindo erros manuais e aumentando a eficiência operacional.

## Uso <a name = "uso"></a>

Estas instruções ajudarão você a configurar e utilizar o projeto para fins de desenvolvimento, teste e operação diária. Siga estes passos para iniciar:

### Pré-requisitos

Você precisará do seguinte para usar os scripts:

- Acesso ao Google Sheets.
- Conta no Google para acesso ao Google Apps Script.
- Permissões adequadas para configurar scripts e utilizar APIs externas (por exemplo, API da Tiny).

### Configuração e Operação

1. **Configuração Inicial:**
   - Acesse o Google Sheets e abra a planilha desejada.
   - No menu, vá para Extensões > Apps Script.
   - Cole o código do script no editor de Apps Script e salve.

2. **Configuração de Gatilhos:**
   - Configure os gatilhos necessários em Extensões > Apps Script > Gatilhos do projeto atual para automatizar a execução dos scripts conforme necessário.

3. **Testes Manuais:**
   - Teste os scripts manualmente para verificar se estão funcionando corretamente.

4. **Operação Diária:**
   - **Atualização Automática:** Configure gatilhos programados para executar os scripts de atualização de estoque e preços automaticamente em intervalos regulares.
   - **Verificação e Log:** Monitore os logs no Google Apps Script para entender o comportamento dos scripts e identificar possíveis erros.
   - **Customização:** Adapte os scripts conforme necessário, ajustando parâmetros de API e manipulação de dados específicos para atender às necessidades da sua empresa.


## Autores <a name = "autores"></a>

- **Gustavo Lima** - *Desenvolvedor inicial* - [gustavo@kamico.com.br](mailto:gustavo@kamico.com.br)
- **Maicon de Menezes** - *Contribuidor* - [maicon@kamico.com.br](mailto:maicon@kamico.com.br)

Para mais detalhes sobre como contribuir para o projeto, consulte o guia em [Contribuição](CONTRIBUTING.md).