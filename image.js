const imghash = require('imghash');
 
imghash
  .hash('sample/2.jpg')
  .then((hash) => {
    console.log(hash); // 'f884c4d8d1193c07'
  });
 
// Custom hex length and result in binary
imghash
  .hash('sample/1.jpg', 4, 'binary')
  .then((hash) => {
    console.log(hash); // '1000100010000010'
  });