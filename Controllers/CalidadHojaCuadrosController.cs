using MultasLectura.Interfaces;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controllers
{
    public class CalidadHojaCuadrosController : ICalidadHojaCuadrosController
    {
        public void CrearTablaDinEmpleadoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinEmpleadoTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["empleado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["compute_0005"]);
            pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum;
        }

        public void CrearTablaDinLectorTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["D1"], rango, "TablaDinLectorTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["lector"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count;
        }
    }
}
