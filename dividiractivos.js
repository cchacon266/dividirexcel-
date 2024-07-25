const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function dividirExcel() {
    try {
        // Ruta del archivo de entrada (absoluta)
        const archivoEntrada = path.join(__dirname, 'activos.xlsx');

        // Directorio de salida (absoluto)
        const carpetaSalida = path.join(__dirname, 'activosDivididos');

        // Crear la carpeta de salida si no existe
        if (!fs.existsSync(carpetaSalida)) {
            fs.mkdirSync(carpetaSalida);
        }

        // Cantidad de filas por archivo
        const cantFilasPorArchivo = 800;

        // Cargar el libro de Excel
        const libro = new ExcelJS.Workbook();
        await libro.xlsx.readFile(archivoEntrada);

        // Seleccionar la hoja de trabajo
        const hojaTrabajo = libro.worksheets[0]; // Cambia a 'worksheets[0]' si quieres la primera hoja, o usa '.getWorksheet('nombre_de_hoja')'

        // Verificar que la hoja de trabajo existe
        if (!hojaTrabajo) {
            throw new Error('La hoja de trabajo no se encontró.');
        }

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

            // Establecer estilos para la fila de encabezado (opcional)
            // nuevaFilaEncabezado.font = filaEncabezado.font;
            // nuevaFilaEncabezado.alignment = filaEncabezado.alignment;

            // Copiar el resto de las filas al nuevo libro
            for (let j = filaInicio; j <= filaFin; j++) {
                const fila = hojaTrabajo.getRow(j);
                nuevaHoja.addRow(fila.values);
            }

            // Guardar el nuevo libro
            const nombreArchivo = `activos_${i}.xlsx`;
            const rutaArchivoSalida = path.join(carpetaSalida, nombreArchivo);
            await nuevoLibro.xlsx.writeFile(rutaArchivoSalida);
        }

        console.log('Archivos divididos correctamente.');
    } catch (error) {
        console.error(error);
    }
}

// Ejecutar la función
dividirExcel().catch(error => console.error(error));