function main(workbook: ExcelScript.Workbook) {
    // Obtener o crear la hoja de consolidado
    let hojaDestino = workbook.getWorksheet("Consolidado");
    if (!hojaDestino) {
        hojaDestino = workbook.addWorksheet("Consolidado");
    }

    // Limpiar la hoja de consolidado antes de agregar datos nuevos
    hojaDestino.getRange().clear();

    // Variables para consolidación
    let filaDestino = 1;
    let valoresUnicos = new Set<string>();
    let datosConsolidados: string[][] = [];

    // Recorrer todas las hojas del libro
    let hojas = workbook.getWorksheets();
    for (let hoja of hojas) {
        if (hoja.getName() !== "Consolidado") {
            let rangoColumnaC = hoja.getRange("C:C").getUsedRange();
            
            if (rangoColumnaC) {
                let valores = rangoColumnaC.getValues();
                
                // Recorrer los valores, omitiendo la primera fila (encabezado)
                for (let i = 1; i < valores.length; i++) { 
                    let valor = valores[i][0];

                    if (valor) {
                        // Agregar a la lista total de valores con el nombre de la hoja
                        datosConsolidados.push([hoja.getName(), valor.toString()]);

                        // Agregar solo valores únicos al Set
                        valoresUnicos.add(valor.toString());
                    }
                }
            }
        }
    }

    // Escribir los valores con el nombre de la hoja en las columnas A y B
    for (let i = 0; i < datosConsolidados.length; i++) {
        hojaDestino.getCell(i, 0).setValue(datosConsolidados[i][0]); // Columna A: Nombre de la hoja
        hojaDestino.getCell(i, 1).setValue(datosConsolidados[i][1]); // Columna B: Valor de la columna C
    }

    // Escribir valores únicos en la columna C
    let filaUnicos = 0;
    valoresUnicos.forEach((valor) => {
        hojaDestino.getCell(filaUnicos, 2).setValue(valor); // Columna C: Valores únicos
        filaUnicos++;
    });

    // Mensaje de finalización
    hojaDestino.getRange("A1").setValue("Consolidación Completa");
}
