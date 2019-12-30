const imghash = require('imghash');

var name1 = 'sample/1.jpg';
var name1_1 = 'sample/1_1.jpg';
var name2 = 'sample/2.jpg'
var name2_1 = 'sample/2_1.jpg'


async function test(){
  var a1 = await imghash.hash(name1)
  await console.log(name1 + ' '+ a1);

var a2 = await imghash.hash(name1_1)
await console.log(name1_1 + ' '+ a2); // 'f884c4d8d1193c07'


var a3 = await imghash.hash(name2)
await console.log(name2 + ' '+ a3);
 
// Custom hex length and result in binary
var a4 = await imghash.hash(name2_1, 4, 'binary')
await console.log(name2_1+' '+a4); // '1000100010000010'
  
}
test();


