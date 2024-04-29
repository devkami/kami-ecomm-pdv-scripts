function KAMI_GPT(input, cell) {
  var apiKey = OPENAI_API_KEY; //substitua OPENAI_API_KEY pelo seu token de API
  var apiUrl = 'https://api.openai.com/v1/chat/completions';
  
  var headers = {
    'Authorization': 'Bearer ' + apiKey,
    'Content-Type': 'application/json'
  };
  
  var message = input + cell + " SEMPRE RESPONDA EM PORTUGUÊS BRASIL";
  
  var payload = {
    model: 'gpt-3.5-turbo',
    messages: [{ role: 'user', content: message }]
  };
  
  var options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload)
  };
  
  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseContent = JSON.parse(response.getContentText());
    var description = responseContent.choices[0].message.content;
    return description;
  } catch (error) {
    Logger.log("Erro ao fazer a solicitação: " + error);
    
    // Aguarde 5 segundos antes de tentar novamente
    Utilities.sleep(62000);
    
    // Chame a função recursivamente para tentar novamente
    return KAMI_GPT(input, cell);
  }
}