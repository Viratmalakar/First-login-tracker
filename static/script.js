function copyTable(){

let table=document.getElementById("tableBox");

html2canvas(table).then(canvas=>{

canvas.toBlob(function(blob){

const item=new ClipboardItem({'image/png': blob});

navigator.clipboard.write([item]);

alert("Table copied as PNG");

});

});

}
