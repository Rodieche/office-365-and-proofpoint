const path = require('path');
const excel = require('./excel');
const validations = require('./validations');

const mailboxes = excel.excelFunction('Exchange.csv');
const users = excel.excelFunction('Users.csv');
const proofpoint = excel.excelFunction('Proofpoint.xlsx');

if(!mailboxes){
  console.log('Exchange.csv file not found'):
  return;
}

if(!users){
  console.log('Users.csv file not found'):
  return;
}

if(!proofpoint){
  console.log('Proofpoint.xlsx file not found'):
  return;
}

let company = null;
let band = 0;

let proofpoint_invalid_list = [];

const mailboxes_check = mailboxes.map(user => {
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
//data_output = validations.validateUserinProofpoint(data_output);

excel.createExcelSheet(company, data_output);
