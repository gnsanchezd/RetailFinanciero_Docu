var request = require('request');

request({
  method: 'GET',
  url: 'https://api.clickup.com/api/v1/team',
  headers: {
    'Authorization': 'pk_A33IPMXOU2UJLFEYWF0WSIYSYC7PTB47'
  }}, function (error, response, body) {
  console.log('Status:', response.statusCode);
  console.log('Headers:', JSON.stringify(response.headers));
  console.log('Response:', body);
});
