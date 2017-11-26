var sqlite3 = require('sqlite3');
let db = new sqlite3.Database('db/storage.db', sqlite3.OPEN_READWRITE, (err) => {
  if (err) {
    console.error(err.message);
  }
  console.log('Connected to the storage.db database.');
});
