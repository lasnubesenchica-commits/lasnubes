function testClaudeApiKey() {
  const key = getClaudeApiKey();
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': key,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 10,
      messages: [{ role: 'user', content: 'di hola' }]
    })
  });
  Logger.log('Status: ' + response.getResponseCode());
  Logger.log('Response: ' + response.getContentText().slice(0, 200));
}