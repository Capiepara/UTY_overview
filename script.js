const fileInput = document.getElementById("excelFile");

const sheetSelect = document.getElementById("sheetSelect");

const fileName = document.getElementById("fileName");

const generateBtn = document.getElementById("generateBtn");

fileInput.addEventListener("change",function(e){

    const file=e.target.files[0];

    if(!file) return;

    fileName.innerText=file.name;

    const reader=new FileReader();

    reader.onload=function(event){

        const data=new Uint8Array(event.target.result);

        const workbook=XLSX.read(data,{type:"array"});

        sheetSelect.innerHTML="";

        workbook.SheetNames.forEach(sheet=>{

            let option=document.createElement("option");

            option.value=sheet;

            option.text=sheet;

            sheetSelect.appendChild(option);

        });

        generateBtn.disabled=false;

    }

    reader.readAsArrayBuffer(file);

});
