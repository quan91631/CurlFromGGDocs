const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const excel = require('exceljs');
const { EACCES } = require('constants');
require('dotenv').config()

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';
// To store the data 
const Arr = []
// Name of the excel file 
const fileName = "Link"

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Docs API.
  authorize(JSON.parse(content), listOfLink);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth 2.0 client.
 */
/** 
 * Print the names and link of students in a spreadsheet:
 */

function listOfLink(auth) {
  const sheets = google.sheets('v4');
  sheets.spreadsheets.values.get(
    {
      auth: auth,
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: 'CLC!A3:M',
    },
      (err, res) => {
        if (err) {
          console.error('The API returned an error.');
          console.log(err);;
        }
        const rows = res.data.values;
        if (rows.length === 0) {
          console.log('No data found.');
        } 
        else {
          for (const row of rows) {
            if (row[4].toLowerCase() == "nữ"){
              const temp = {name: row[1], link: row[12]}
              Arr.push(temp)
            }
          }
            insertExcel(Arr)      
        }
      })
    }       
async function insertExcel(newArr){
  let workBook = new excel.Workbook()
  let workSheet = workBook.addWorksheet('Link')

  // setting the columns key

  workSheet.columns = [
    {header: 'Name', key: 'name',width: 32 }, 
    {header:'Link', key:'link',width: 32}
  ];

  // setting Column width is at least the header width

  workSheet.columns.forEach(column => {
    column.width = column.header.length < 12 ? 12 : column.header.length
  })

  // Adding the data into each row base on the key  
  newArr.forEach((item,index) =>{
    let row = workSheet.getRow(index + 1);
      row.values = item;
    })
    await workBook.xlsx.writeFile(`${fileName}.xlsx`)
}

