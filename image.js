const imghash = require('imghash');

var name1 = 'sample/1.jpg';
var name1_1 = 'sample/1_1.jpg';
var name2 = 'sample/2.jpg'
var name2_1 = 'sample/2_1.jpg'


async function test(){
  imghash
  .hash(name1)
    await console.log(name1 + ' '+ hash);

imghash
  .hash(name1_1)
    await console.log(name1_1 + ' '+ hash); // 'f884c4d8d1193c07'


imghash
  .hash(name2)
    await console.log(name2 + ' '+ hash);
 
// Custom hex length and result in binary
imghash
  .hash(name2_1, 4, 'binary')
    await console.log(name2_1+' '+hash); // '1000100010000010'
  
}
test();


