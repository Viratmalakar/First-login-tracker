function filterTable(){

let input=document.getElementById("agentFilter").value.toLowerCase();

let from=document.getElementById("fromDate").value;

let to=document.getElementById("toDate").value;

let table=document.getElementById("dataTable");

let rows=table.getElementsByTagName("tr");

for(let i=1;i<rows.length;i++){

let id=rows[i].cells[0].innerText.toLowerCase();
let name=rows[i].cells[1].innerText.toLowerCase();

if(id.includes(input)||name.includes(input))
rows[i].style.display="";
else
rows[i].style.display="none";

}

if(from!=""&&to!=""){

let cells=document.querySelectorAll("td[data-date]");

cells.forEach(cell=>{

let d=cell.dataset.date;

if(d<from||d>to)
cell.style.display="none";
else
cell.style.display="";

});

}

}


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
