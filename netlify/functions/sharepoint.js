const https = require('https');
const querystring = require('querystring');

const TENANT_ID = "0253ea9c-ecdf-46be-837f-73ea9cbaad44";
const CLIENT_ID = "6c023227-33fe-4358-b50a-4b6f6d6213e8";
const CLIENT_SECRET = "19d8Q~DrkwSYNEUzHV8inBIwXbtBy8p~QjkJxc9v";

const H = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Content-Type": "application/json"
};

function post(hostname, path, bodyObj) {
  return new Promise((resolve, reject) => {
    const body = querystring.stringify(bodyObj);
    const req = https.request({
      hostname, path, method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString();
        try { resolve(JSON.parse(raw)); }
        catch(e) { reject(new Error('Parse error: ' + raw.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: H, body: '' };
  try {
    const result = await post(
      'login.microsoftonline.com',
      `/${TENANT_ID}/oauth2/v2.0/token`,
      { grant_type:'client_credentials', client_id:CLIENT_ID, client_secret:CLIENT_SECRET, scope:'https://graph.microsoft.com/.default' }
    );
    if (!result.access_token) return { statusCode:400, headers:H, body:JSON.stringify({error:'No token',details:result}) };
    return { statusCode:200, headers:H, body:JSON.stringify({access_token:result.access_token}) };
  } catch(e) {
    return { statusCode:500, headers:H, body:JSON.stringify({error:e.message}) };
  }
};
