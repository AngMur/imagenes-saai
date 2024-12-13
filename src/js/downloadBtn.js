// Objeto que almacena la clave correspondiente a cada carrera. 
// Las claves son códigos únicos asociados a cada nombre de carrera.
const carreraClave = {
    "Arquitectura": "102",
    "Diseño industrial": "105",
    "Ingenieria civil": "107",
    "Ingenieria Electrica Electronica": "109",
    "Ingenieria en computacion": "110",
    "Ingenieria Industrial": "114",
    "Ingenieria Mecanica": "115",
    "Derecho": "305",
    "Derecho (SUA)": "305SUA",
    "Economia": "306",
    "Economia(SUA)": "306SUA",
    "Desarrollo Agropecuario": "309",
    "Relaciones Internacionales": "310",
    "Relaciones Internacionales (SUA)": "310SUA",
    "Sociologia": "311",
    "Comunicacion y periodismo": "316",
    "Pedagogia": "421"
};

// Arreglo que contiene las materias. Estas se utilizan para asociarlas cíclicamente con las hojas del archivo Excel.
const materias = [
    "L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9"
];

/**
 * Función principal para procesar un archivo Excel y generar un archivo ZIP con imágenes.
 * @param {File} file - Archivo Excel cargado por el usuario.
 * @param {string} ceilRange - Rango de celdas a extraer, por ejemplo, "A1:M125".
 * @param {boolean} isNotDiagnostic - Indica si el archivo corresponde a diagnósticos o no.
 * @param {string} signature - Firma o identificador de la materia (opcional para diagnósticos).
 */
function download(file, ceilRange, isNotDiagnostic = false, signature = '') {
    const reader = new FileReader();
    reader.onload = function (event) {
        const data = event.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });

        const zip = new JSZip(); // Crear archivo ZIP para almacenar las imágenes

        // Iterar por todas las hojas del archivo Excel
        for (let i = 0; i < workbook.SheetNames.length; i++) {
            if (i === 0) {
                continue; // Ignorar la primera hoja
            }

            const sheetName = workbook.SheetNames[i];
            const sheet = workbook.Sheets[sheetName];

            // Obtener los datos del rango especificado
            const range = XLSX.utils.decode_range(ceilRange);
            const dataRange = [];
            for (let row = range.s.r; row <= range.e.r; row++) {
                let rowData = [];
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = { r: row, c: col };
                    const cell = sheet[XLSX.utils.encode_cell(cellAddress)];
                    rowData.push(cell ? cell.v : "");
                }
                dataRange.push(rowData);
            }

            // Obtener la columna con valor maximo (omitiendo primera y última fila)
            const secondColumnValues = dataRange.slice(1, -1).map(row => row[1]); 
            const maxSecondColumn = Math.max(...secondColumnValues.filter(val => typeof val === 'number'));

            // Identificar la fila con el valor máximo en la segunda columna
            const maxRowIndex = secondColumnValues.findIndex(val => val === maxSecondColumn) + 1;

            // Determinar la clave de carrera y la materia
            const carrera = isNotDiagnostic 
                ? Object.values(carreraClave)[Math.floor((i - 1))] 
                : Object.values(carreraClave)[Math.floor((i - 1) / 9)];
            const materia = isNotDiagnostic 
                ? signature 
                : materias[(i - 1) % 9];
            console.log(carrera, materia, workbook.SheetNames[i]);

            // Crear la tabla para los datos extraídos
            const container = document.createElement('div');
            container.style.position = 'absolute';
            document.body.appendChild(container);

            const table = document.createElement('table');
            table.className = 'styled-table';

            const tbody = document.createElement('tbody');
            dataRange.forEach((row, index) => {
                const tableRow = document.createElement('tr');

                // Resaltar la fila con el valor máximo
                if (index === maxRowIndex) {
                    const colors = ["#4bfa4b", "#4bfa4b", "#ffeb3b", '#ff6363', "#ff6363"];
                    tableRow.style.backgroundColor = colors[maxRowIndex - 1];
                }

                row.forEach(cell => {
                    const cellElement = document.createElement('td');
                    // Formatear números con dos decimales
                    cellElement.textContent = (typeof cell === 'number') ? cell.toFixed(2) : cell;
                    tableRow.appendChild(cellElement);
                });
                tbody.appendChild(tableRow);
            });

            table.appendChild(tbody);
            container.appendChild(table);

            // Convertir tabla en imagen y añadirla al archivo ZIP
            html2canvas(container).then(function (canvasImg) {
                const imgData = canvasImg.toDataURL('image/png');
                zip.file(`${carrera}_${materia}_2026.png`, imgData.split(',')[1], { base64: true }); 
                document.body.removeChild(container);

                // Generar el ZIP y descargarlo al final
                if (i === workbook.SheetNames.length - 1) {
                    zip.generateAsync({ type: "blob" }).then(function (content) {
                        const link = document.createElement('a');
                        link.href = URL.createObjectURL(content);
                        link.download = isNotDiagnostic 
                            ? (signature === 'L10' ? "Español.zip" : "Ingles.zip") 
                            : "Diagnostico.zip";
                        link.click();
                    });
                }
            });
        }
    };
    reader.readAsBinaryString(file);
}
