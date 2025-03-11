function main(workbook: ExcelScript.Workbook) {
    // Obtener o crear la hoja de consolidado
    let hojaDestino = workbook.getWorksheet("Consolidado");
    if (!hojaDestino) {
        hojaDestino = workbook.addWorksheet("Consolidado");
    }

    // Limpiar la hoja de consolidado antes de agregar datos nuevos
    hojaDestino.getRange().clear();

    let filaDestino = 1;
    let valoresUnicos = new Set<string>();
    let valoresTotales: string[] = [];

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
                        // Agregar a la lista total de valores
                        valoresTotales.push(valor.toString());

                        // Agregar solo valores únicos al Set
                        valoresUnicos.add(valor.toString());
                    }
                }
            }
        }
    }

    // Escribir todos los valores en la columna A
    for (let i = 0; i < valoresTotales.length; i++) {
        hojaDestino.getCell(i, 0).setValue(valoresTotales[i]); // Columna A
    }

    // Escribir valores únicos en la columna B
    let filaUnicos = 0;
    valoresUnicos.forEach((valor) => {
        hojaDestino.getCell(filaUnicos, 1).setValue(valor); // Columna B
        filaUnicos++;
    });

    // Mensaje de finalización
    hojaDestino.getRange("A1").setValue("Consolidación Completa");
}
