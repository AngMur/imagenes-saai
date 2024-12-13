const dropArea1 = document.getElementById("dp1");
const dropArea2 = document.getElementById("dp2");
const dropArea3 = document.getElementById("dp3");

const inputFile1 = document.getElementById("input-1");
const inputFile2 = document.getElementById("input-2");
const inputFile3 = document.getElementById("input-3");

/* Cada input file asi como su drop area se inicializa para que al ser arrastrado un archivo o al ser seleccionado se mande a llamar la funcion que se encarga de generar las imagenes la funcion (download)*/;

// INIT FILE 1
inputFile1.addEventListener("change", uploadFile1);

dropArea1.addEventListener("dragover", function(e) {
    e.preventDefault();
});

dropArea1.addEventListener("drop", function(e) {
    e.preventDefault();
    let files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        inputFile1.files = files;
        uploadFile1();
    } else {
        alert("Por favor, sube solo archivos excel");
    }
});

// Se manda a llamar la función que lee el archivo y genera las imagenes 
function uploadFile1() {
    //La función download puede recibir 4 parametros
    //(archivoALeer, RangoALeer, BanderaParaIndicarQueNoEsDiagnostico, MateriaAColocarEnNomenclatura)
    download(inputFile1.files[0], 'O4:U10');
    inputFile1.value = "";
}

// INIT FILE 2
inputFile2.addEventListener("change", uploadFile2);

dropArea2.addEventListener("dragover", function(e) {
    e.preventDefault();
});

dropArea2.addEventListener("drop", function(e) {
    e.preventDefault();
    let files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        inputFile2.files = files;
        uploadFile2();
    } else {
        alert("Por favor, sube solo archivos excel");
    }
});

function uploadFile2() {
    download(inputFile2.files[0], 'W6:AC12', true, 'L10'); // Cambia el rango según lo necesario
    inputFile2.value = "";
}

// INIT FILE 3
inputFile3.addEventListener("change", uploadFile3);

dropArea3.addEventListener("dragover", function(e) {
    e.preventDefault();
});

dropArea3.addEventListener("drop", function(e) {
    e.preventDefault();
    let files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        inputFile3.files = files;
        uploadFile3();
    } else {
        alert("Por favor, sube solo archivos excel");
    }
});

function uploadFile3() {
    download(inputFile3.files[0], 'V6:AB11', true, 'L11'); // Cambia el rango según lo necesario
    inputFile3.value = "";
}
