using MultasLectura.LibroPlazos.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroPlazos.Controllers
{
    public class PlazosHojaResumenController : IPlazosHojaResumenController
    {
        public void CrearTablaDatosTarifa(ExcelWorksheet hoja)
        {
            int numFilaInicial = 3;

            hoja.Cells[$"A1"].Value = "DIAS HÁBILES T1";

            hoja.Cells["A1:C1"].Merge = true;
            hoja.Cells[$"A2"].Value = "FTL";
            hoja.Cells[$"B2"].Value = "k";
            hoja.Cells[$"C2"].Value = "Qij";

            List<int> ftl = new()
          ;

            List<int> k = new();

            for(int i = -14; i < 18; i++)
            {
                ftl.Add(i);

            }

            foreach(int num in ftl)
            {
                hoja.Cells[$"A{numFilaInicial}"].Value = num;

                numFilaInicial++;

            }



         
        }

        public void CrearTablaDinCertAtrasoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["Z1"], rango, "TablaDinamicaCertAtrasoTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["tip_itin"]);
            var atraso = pivotTable.RowFields.Add(pivotTable.Fields["atraso"]);
            pivotTable.DataFields.Add(pivotTable.Fields["cant_sum"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Sum;

            atraso.Sort = eSortType.Ascending;
        }

        public void CrearTablaImportesFinales()
        {
            throw new NotImplementedException();
        }
    }
}
