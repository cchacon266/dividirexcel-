const ExcelJS = require('exceljs');
const path = require('path');

async function dividirExcel() {
    // Ruta del archivo de entrada (absoluta)
    const archivoEntrada = path.join(__dirname, 'referencias.xlsx');

    // Directorio de salida (absoluto)
    const carpetaSalida = path.join(__dirname, 'referenciasDivididos');

    // Cantidad de filas por archivo
    const cantFilasPorArchivo = 800;

    // Cargar el libro de Excel
    const libro = new ExcelJS.Workbook();
    await libro.xlsx.readFile(archivoEntrada);

    // Seleccionar la hoja de trabajo
    const hojaTrabajo = libro.getWorksheet(1);

    // Calcular la cantidad total de archivos necesarios
    const ultimaFila = hojaTrabajo.rowCount;
    const cantArchivos = Math.ceil(ultimaFila / cantFilasPorArchivo);

    // Crear archivos divididos
    for (let i = 1; i <= cantArchivos; i++) {
        const filaInicio = (i - 1) * cantFilasPorArchivo + 1;
        const filaFin = Math.min(i * cantFilasPorArchivo, ultimaFila);

        // Crear un nuevo libro
        const nuevoLibro = new ExcelJS.Workbook();
        const nuevaHoja = nuevoLibro.addWorksheet('Sheet 1');

        // Copiar fila de encabezado al nuevo libro
        const filaEncabezado = hojaTrabajo.getRow(1);
        const nuevaFilaEncabezado = nuevaHoja.addRow(filaEncabezado.values);
        
        // Establecer estilos para la fila de encabezado
        nuevaFilaEncabezado.font = filaEncabezado.font;
        nuevaFilaEncabezado.alignment = filaEncabezado.alignment;

        // Copiar el resto de las filas al nuevo libro
        for (let j = filaInicio; j <= filaFin; j++) {
            const fila = hojaTrabajo.getRow(j);
            nuevaHoja.addRow(fila.values);
        }

        // Guardar el nuevo libro
        const nombreArchivo = `references_${i}.xlsx`;
        const rutaArchivoSalida = path.join(carpetaSalida, nombreArchivo);
        await nuevoLibro.xlsx.writeFile(rutaArchivoSalida);
    }

    console.log('Archivos divididos correctamente.');
}

// Ejecutar la función
dividirExcel().catch(error => console.error(error));
