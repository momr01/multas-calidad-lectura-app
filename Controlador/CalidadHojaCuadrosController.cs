using MultasLectura.Controlador.Interfaz;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controlador
{
    public class CalidadHojaCuadrosController : ICalidadHojaCuadrosController
    {
        public void CrearTablaDinEmpleadoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            // Crear tabla dinámica
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinEmpleadoTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["empleado"]);
            //pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["compute_0005"]);
            pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum;

            //label2.Text =  pivotTable.Fields.Count.ToString();
        }

        public void CrearTablaDinLectorTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            // Crear tabla dinámica
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["D1"], rango, "TablaDinLectorTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["lector"]);
            //pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count;

            //label2.Text =  pivotTable.Fields.Count.ToString();
        }
    }
}
