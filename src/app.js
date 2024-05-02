const path = require('path');
const excel = require('./excel');
const consoleEmojis = require('console-emojis');
const colors = require('colors');


const validations = require('./validations');

const mailboxes = excel.excelFunction('Exchange.csv');
const users = excel.excelFunction('Users.csv');
const proofpoint = excel.excelFunction('Proofpoint.xlsx');

if(!mailboxes){
  console.x('Exchange.csv file not found');
  return;
}else{
  console.ok('Exchange.csv file found')
}

if(!users){
  console.x('Users.csv file not found');
  return;
}else{
  console.ok('Users.csv file found')
}

if(!proofpoint){
  console.x('Proofpoint.xlsx file not found');
  return;
}else{
  console.ok('Proofpoint.xlsx file found')
}

console.log('Analyzing data ...')

let company = null;
let band = 0;

let proofpoint_invalid_list = [];

const mailboxes_check = mailboxes.map(user => {
  console.question('Analyzing', user['Email address'].toLowerCase().green, 'information...' );
  const licenses = users.filter(u => u['User principal name'].toLowerCase() === user['Email address'].toLowerCase())[0]?.['Licenses'];
  let proofpoint_valid = validations.filterObjectsBySubstring(proofpoint, 'Email', user['Email address'] )[0]?.Role || undefined;

  if(band === 0){
    company = user['Email address'].split('@')[1].split('.')[0];
    band = 1;
  }

  const dataUser = {
    'Company': company,
    'Display Name': user['Display name'],
    'Email': user['Email address'],
    'Recipient Type': user['Recipient type'],
    'Licenses': licenses?.replace('+', ',') ?? 'Unlicensed',
    'Proofpoint': proofpoint_valid ? proofpoint_valid : 'No Applied',
    'Observations': ''
  }

  dataUser.Observations = validations.validateUserinProofpoint(dataUser);
  return dataUser;
});

proofpoint.forEach(function(up){
    let invalid = true;
    const email = up.Email.toLowerCase();
    mailboxes.forEach(function(um){
        const mailbox_email = um['Email address'].toLowerCase();
        if(email.includes(mailbox_email)) invalid = false;
    });
    if (invalid){
        proofpoint_invalid_list.push(up)
    }
});

const proofpoint_invalids = proofpoint_invalid_list.map(up => {
  console.question('Analyzing', up['Email'].toLowerCase().gray, 'information...' );
    return {
        'Company': company,
        'Display Name': 'N/A',
        'Email': up.Email,
        'Recipient Type': 'N/A',
        'Licenses': 'N/A',
        'Proofpoint': up.Role,
        'Observations': 'PROOFPOINT: Invalid user in Proofpoint'
      }
})

let data_output = [...mailboxes_check, ...proofpoint_invalids];

excel.createExcelSheet(company, data_output);

console.ok_hand('Process finished'.yellow)
