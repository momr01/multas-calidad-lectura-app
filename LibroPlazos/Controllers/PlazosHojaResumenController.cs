using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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
        public void CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles)
        {
            int numFilaInicial = 3;
            int numFila = 3;

            hojaResumen.Cells[$"A1"].Value = "DIAS HÁBILES T1";

            hojaResumen.Cells["A1:C1"].Merge = true;
            hojaResumen.Cells[$"A2"].Value = "FTL";
            hojaResumen.Cells[$"B2"].Value = "k";
            hojaResumen.Cells[$"C2"].Value = "Qij";

            List<int> ftl = new()
          ;

            List<int> k = new();

            for(int i = -14; i < 18; i++)
            {
                ftl.Add(i);

            }

            foreach(int num in ftl)
            {
                hojaResumen.Cells[$"A{numFila}"].Value = num;

                int cantidad = CalcularCantAtrasos(hojaReclDetalles, num);

                hojaResumen.Cells[$"C{numFila}"].Value = cantidad;

                ColorearRangoSegunNum(num, hojaResumen, numFila);

                if(numFila == numFilaInicial)
                {
                    hojaResumen.Cells[$"A{numFila}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                    hojaResumen.Cells[$"A{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                    hojaResumen.Cells[$"A{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                }


                numFila++;

            }

          // LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"A{numFilaInicial}:A{numFila}"]);




        }

        private void ColorearRangoSegunNum(int num, ExcelWorksheet hojaResumen, int numFilaInicial)
        {
            if (num <= -6 || num >= 4)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"A{numFilaInicial}:C{numFilaInicial}"], Color.FromArgb(1, 255, 102, 0));
            }
            else if (num == -5 || num == 3)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"A{numFilaInicial}:C{numFilaInicial}"], Color.FromArgb(1, 255, 204, 153));
            }
            else if (num == -4 || num == 2)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"A{numFilaInicial}:C{numFilaInicial}"], Color.FromArgb(1, 255, 255, 153));
            }
            else
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"A{numFilaInicial}:C{numFilaInicial}"], Color.FromArgb(1, 204, 255, 204));
            }
            /* List<int> ftl = new()
          ;

             for (int i = -14; i < 18; i++)
             {
                 ftl.Add(i);

             }

             foreach (int num in ftl)
             {
                if(num <= -6 || num >= 4)
                 {
                     LibroExcelHelper.FondoSolido();
                 }

             }*/



        }



        private int CalcularCantAtrasos(ExcelWorksheet hojaBase, int atraso)
        {
            int cantFilas = hojaBase.Dimension.Rows;
          

            int colTipItin = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "tip_itin");
            int colAtraso = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "atraso");
            int colCantSum = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "cant_sum");

            int cantFinal = 0;

            if (colTipItin != -1 && colAtraso != -1 && colCantSum != -1)
            {
                for (int fila = 1; fila <= cantFilas; fila++)
                {
                    object cellValue = hojaBase.Cells[fila, colTipItin].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.ToLower().Contains("itinerario t1"))
                        {
                            int atrasoBase = int.Parse(hojaBase.Cells[fila, colAtraso].Value.ToString()!);

                            if(atrasoBase == atraso)
                            {
                                cantFinal += int.Parse(hojaBase.Cells[fila, colCantSum].Value.ToString()!);
                            }
                           


                        }
                    }
                }
            }

            return cantFinal;
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
