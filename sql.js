'use strict';

const sqlite3 = require('sqlite3').verbose();
 
let db = new sqlite3.Database('./db/sample.db');
 
let sql = `select * from viitenumbrid where viitenum is not null limit 1`;


function parser (query){

    db.each(query, function(err, row) {
        let referenceNumber = (row.viitenum);
              return referenceNumber;
    });
	

};

let result = parser(sql);
console.log(result);

//Код выше, выдает undefined. Т.е. как я понял, значение переменной referenceNumber не вышло за пределы функции parser. Вопрос в том, как мне вернуть результаты запроса и присвоить их переменной result 



