var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
//Load the docx file as a binary
var content = fs.readFileSync(path.resolve('BlankDoc.docx'), 'binary');
var zip = new JSZip(content);
var doc = new Docxtemplater();
doc.loadZip(zip);
//set the templateVariables

function generateFile() {
    var nimi = document.getElementById("mainForm").elements[0].value;
    var kood = document.getElementById("mainForm").elements[1].value;
    var elukoht = document.getElementById("mainForm").elements[2].value;
     var mark = document.getElementById("mainForm").elements[3].value;
     var kuupaev = document.getElementById("mainForm").elements[4].value;
     var aeg = document.getElementById("mainForm").elements[5].value;
     var  koht = document.getElementById("mainForm").elements[6].value;
     var kirjeldus = document.getElementById("mainForm").elements[7].value;
    var Today = new Date();
    var  koostamiseKPV=Today.getDate() + "." + (Today.getMonth()+1)  + "." + Today.getFullYear();
    var kvalifikatsioon="seadus"
doc.setData({
    nimi: nimi,
    kood: kood,
    elukoht: elukoht,
    mark: mark,
    kuupaev: kuupaev,
    aeg: aeg,
    koht: koht,
    kirjeldus: kirjeldus,
    kvalifikatsioon: kvalifikatsioon,
    koostamiseKPV: koostamiseKPV
});

try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
}
catch (error) {
    var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties,
    }
    console.log(JSON.stringify({error: e}));
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
    throw error;
}

var buf = doc.getZip().generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(mark+'.docx'), buf);

//Writing to database
var sqlite3 = require('sqlite3').verbose();
var db = new sqlite3.Database('./db/sample.db');

 db.run(`INSERT INTO langs(name) VALUES(?)`, [nimi]);
  db.close();

};

//delete from viitenumbrid where viitenum in (select viitenum from viitenumbrid order by viitenum limit 1)