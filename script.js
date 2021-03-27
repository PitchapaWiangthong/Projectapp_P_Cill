var xlsx = require("xlsx");
var wb = xlsx.readFile("Word.xlsx",{cellDates:true});
var ws = wb.Sheets["Sheet1"];
var maxRows = xlsx.utils.sheet_to_row_object_array(ws).length;
var min = 1;
var i = 0;
var array = [];

// while (i < maxRows) {
    // var  s    = random(min,maxRows,array);
    // console.log(s)
//     i++;
// }
console.log(random(min,maxRows,array));


/////////////////////////////////////////////////////////////////////////
function random(min,max,array) {
    var i = Math.floor(Math.random()*(max-min+1))+min;
    for (let j = 0 ; j < array.length ; ){
        if(i === array[j]){
            i = Math.floor(Math.random()*(max-min+1))+min;
            j = 0;
        }
        else{
            j++;
        }
    }
    array.push(i);
    return i;
}
/////////////////////////////////////////////////////////////////////////

//  function word(idx) {
//      return xlsx.utils.sheet_to_row_object_array(ws)[idx].Words;
//  }
//  function spell(idx) {
//      return xlsx.utils.sheet_to_row_object_array(ws)[idx].Spelling;
//  }

 