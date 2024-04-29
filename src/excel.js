const reader = require('xlsx');
const path = require('path');

const excelFunction = (name) => {

    const mypath = path.join(process.cwd(), 'src', 'raw_data', name); 

    const file = reader.readFile(mypath);

    let data = [];

    const sheets = file.SheetNames 
    
    for(let i = 0; i < sheets.length; i++) 
    { 
    const temp = reader.utils.sheet_to_json( 
        file.Sheets[file.SheetNames[i]]) 
    temp.forEach((res) => { 
        data.push(res) 
    }) 
    } 
    
    return data;
}

const createExcelSheet = (companyName, data) => {
    const newBook = reader.utils.book_new();
    const newSheet = reader.utils.json_to_sheet(data);
    reader.utils.book_append_sheet(newBook, newSheet, "Digest");
    reader.writeFile(newBook, path.join(process.cwd(), 'output', `${companyName}_digest.xlsx`) ) 
    return;
}

const updateExcelSheet = (data) => {
    const mypath = path.join(process.cwd(), 'output', 'digest.xlsx'); 
    const workbook = reader.readFile(mypath);
    reader.utils.sheet_add_json(workbook.Sheets["Sheet1"], data);
    reader.writeFile(workbook, mypath)
}


module.exports = {
    excelFunction,
    createExcelSheet,
    updateExcelSheet
}