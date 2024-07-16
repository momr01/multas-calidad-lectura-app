using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Services
{
    public class CalidadHojaCuadrosService
    {
        public void TablaDinamica(ExcelWorksheet hoja,
           string primeraCelda,
           ExcelRange rango,
           string nombreTabla,
           string headerFilas,
           string headerValores,
           OfficeOpenXml.Table.PivotTable.DataFieldFunctions operacion
           )
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells[primeraCelda], rango, nombreTabla);
            pivotTable.RowFields.Add(pivotTable.Fields[headerFilas]);
            pivotTable.DataFields.Add(pivotTable.Fields[headerValores]);
            pivotTable.DataFields[0].Function = operacion;
        }
    }
}
