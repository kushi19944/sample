const imghash = require('imghash');

var name1 = 'sample/1.jpg';
var name1_1 = 'sample/1_1.jpg';
var name2 = 'sample/2.jpg'
var name2_1 = 'sample/2_1.jpg'


async function test(){
  imghash
  .hash(name1)
  .then((hash) => {
    await console.log(name1 + ' '+ hash); // 'f884c4d8d1193c07'
  });

imghash
  .hash(name1_1)
  .then((hash) => {
    await console.log(name1_1 + ' '+ hash); // 'f884c4d8d1193c07'
  });


imghash
  .hash(name2)
  .then((hash) => {
    await console.log(name2 + ' '+ hash); // 'f884c4d8d1193c07'
  });
 
// Custom hex length and result in binary
imghash
  .hash(name2_1, 4, 'binary')
  .then((hash) => {
    await console.log(name2_1+' '+hash); // '1000100010000010'
  });


await Promise
  .all([name1, name1_1])
  .then((results) => {
    const dist = leven(results[0], results[1]);
    console.log(`Distance between images is: ${dist}`);
    if (dist <= 12) {
      console.log('Images are similar');
    } else {
      console.log('Images are NOT similar');
    }
  });


await Promise
  .all([name2, name2_1])
  .then((results) => {
    const dist = leven(results[0], results[1]);
    console.log(`Distance between images is: ${dist}`);
    if (dist <= 12) {
      console.log('Images are similar');
    } else {
      console.log('Images are NOT similar');
    }
}
test();


