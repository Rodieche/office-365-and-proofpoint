const filterObjectsBySubstring = (objects, key, substring) => {
  return objects.filter(obj =>
    obj.hasOwnProperty(key) && obj[key].toString().toLowerCase().includes(substring.toLowerCase())
  );
}

const validateUserinProofpoint = (data) => {
    let comment = '';
    if (
      data['Recipient Type'] === 'UserMailbox' &&
      data['Proofpoint'] === 'No Applied'
    ){
      comment += 'PROOFPOINT: Account needs End User license'
    }else if(
      data['Recipient Type'] === 'UserMailbox' &&
      data['Proofpoint'] !== 'End User' &&
      data['Proofpoint'] !== 'Silent User' &&
      data['Proofpoint'] !== 'Organization Admin'
    ){
      comment += '\n PROOFPOINT: User account applied as Shared Mailbox'
      // du['Observations'] += comment;
    }

    if (
      data['Recipient Type'] === 'SharedMailbox' &&
      data['Proofpoint'] === 'No Applied'
    ){
      comment += 'PROOFPOINT: Account needs a shared mailbox license'
    }else if(
      data['Recipient Type'] === 'SharedMailbox' &&
      data['Proofpoint'] !== 'Distribution Group' &&
      data['Proofpoint'] !== 'Security Group'
    ){
      comment += '\n PROOFPOINT: Shared account applied as End User mailbox'
    }

    if (
      data['Recipient Type'] === 'SharedMailbox' &&
      data['Licenses'] !== 'Unlicensed'
    ){
      comment += '\n OFFICE 365: Shared Mailbox with license applied'
    }


    return comment.trim()
  
}

module.exports = {
    filterObjectsBySubstring,
    validateUserinProofpoint
}

