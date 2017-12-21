var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
//Load the docx file as a binary
var content = fs.readFileSync(path.resolve('BlankDoc.docx'), 'binary');
var zip = new JSZip(content);
var doc = new Docxtemplater();
doc.loadZip(zip);
var sqlite3 = require('sqlite3').verbose();
var db = new sqlite3.Database('./db/sample.db');

//set the templateVariables
function clearForm(){
    document.getElementById("mainForm").reset(); 
};

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
    switch (kirjeldus) {
        case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 361 (peatumise keeld) mõjupiirkonnas.':
        var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
        case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 362 (parkimise keeld) mõjupiirkonnas.':
        var   norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 383 (peatumise keelu ala) mõjupiirkonnas.':
         var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
          case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 384 (parkimise keelu ala) mõjupiirkonnas.':
         var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 363 (parkimiskeeld paaritul kuupäeval) mõjupiirkonnas.':
        var   norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 364 (parkimiskeeld paaris kuupäeval) mõjupiirkonnas.':
         var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - liiklusmärgi nr 874 (puudega inimese sõiduk) nõudeid eirates. Mootorsõidukis puudus puudega inimese sõiduki parkimiskaart':
         var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - teekattemärgise nr 976 (puudega inimese sõiduki parkimiskoht) nõudeid eirates.':
         var  norm = 'liiklusseadus § 21 lg 4 p 2';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - teekattemärgise nr 931 (peatamiskeelu joon) nõudeid eirates.':
         var  norm = 'liiklusseadus § 21 lg 4 p 3';
        break;
         case 'mootorsõiduk pargitud kohas, kus liikluskorraldusvahend seda ei luba, aga nimelt - teekattemärgise nr 932 (parkimiskeelu joon) nõudeid eirates.':
         var  norm = 'liiklusseadus § 21 lg 4 p 4';
        break;
         case 'mootorsõiduk pargitud keelatud kohas - osaliselt või täielikult kõnniteel, kus ei olnud liikluskorraldusvahenditega ettenähtud parkimist.':
         var  norm = 'liiklusseadus § 21 lg 4 p 3';
        break;
         case 'mootorsõiduk pargitud keelatud kohas - ülekäigurajal.':
       var    norm = 'liiklusseadus § 21 lg 2 p 6';
        break;
         case 'mootorsõiduk pargitud keelatud kohas - lähemal kui 5m enne jalakäijate ülekäigurada.':
       var    norm = 'liiklusseadus § 21 lg 2 p 6';
        break;
         case 'mootorsõiduk pargitud keelatud kohas, kus sõidurada (suunavööndit) tähistava pideva märgisjoone ja peatatud sõiduki vahe oli alla 3m.':
       var    norm = 'liiklusseadus § 21 lg 2 p 7';
        break;
         case 'mootorsõiduk pargitud keelatud kohas - haljasalal ilma selle omaniku (valdaja) loata.':
       var    norm = 'liiklusseadus § 21 lg 2 p 12';
        break;
         case 'mootorsõiduk pargitud keelatud kohas - teekattele märgitud parkimiskohtade kõrval.':
        var   norm = 'liiklusseadus § 21 lg 4 p 4';
        break;
         case 'mootorsõiduk pargitud kohas, kus takistas teiste sõidukite lahkumise parkimiskohtadelt.':
       var    norm = 'liiklusseadus § 21 lg 4 p 8';
        break;
         case 'mootorsõiduk pargitud viisil, mis takistab välja-/sissesõitu hoovi/teega külgnevale alale/ garaaži, õuealale':
       var    norm = 'liiklusseadus § 20 lg 1';
        break;
         
          default:
    }
    
doc.setData({
    nimi: nimi,
    kood: kood,
    elukoht: elukoht,
    mark: mark,
    kuupaev: kuupaev,
    aeg: aeg,
    koht: koht,
    kirjeldus: kirjeldus,
    norm: norm,
    koostamiseKPV: koostamiseKPV,
    viitenumber: viitenumber
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