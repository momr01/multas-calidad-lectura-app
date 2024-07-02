using MultasLectura.Controlador.Interfaz;
using MultasLectura.Modelo;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controlador
{
    public class CalidadHojaResumenController : ICalidadHojaResumenController
    {
        private BaremoModel baremos;
        public CalidadHojaResumenController()
        {
            baremos = new BaremoModel
            {
                T1 = 304.91,
                T2 = 2814.51,
                T3 = 304.91,
                AlturaT1 = 3572.63,
                AlturaT3 = 3572.63
            };
        }

        public void CrearTablaBaremosMetas(ExcelWorksheet hoja)
        {
            hoja.Cells["F1"].Value = "Baremo Lectura desde el 01/02/2024";
            hoja.Cells["F2"].Value = "T1 y T3";
            hoja.Cells["F3"].Value = "T2";
            hoja.Cells["F4"].Value = "Altura T1 y T3";
            hoja.Cells["F5"].Value = "Meta";
            hoja.Cells["F6"].Value = "Meta 2";
            hoja.Cells["F7"].Value = "Obtenido";

            LibroExcelModel.AplicarBordesARango(hoja.Cells["F2:G7"]);
        }

        public void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango)
        {
            // Crear tabla dinámica
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinTipoEstado");
            pivotTable.RowFields.Add(pivotTable.Fields["tipo_certificacion"]);
            pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count;

            //label2.Text =  pivotTable.Fields.Count.ToString();

           
        }

        public void CrearTablaMetodoLineal(ExcelWorksheet hojaDestino, ExcelWorksheet hojaOrigen)
        {
            


            // Obtener el número total de filas y columnas en la hoja de cálculo
            int rowCount = hojaOrigen.Dimension.Rows;
            int colCount = hojaOrigen.Dimension.Columns;

            //var celdaA25 = hojaResumen.Cells["A25"];
            hojaDestino.Cells["A25"].Value = "Método Lineal";
            //celdaA25.Value = "Método Lineal";
            hojaDestino.Cells["A26"].Value = "Certificación Itinerario T1";
            hojaDestino.Cells["A27"].Value = "Certificación Itinerario T2";
            hojaDestino.Cells["A28"].Value = "Certificación Itinerario T3";
            hojaDestino.Cells["A29"].Value = "Certificación Itinerario en Altura T1";
            hojaDestino.Cells["A30"].Value = "Certificación Itinerario en Altura T3";

            // ApplyBorders(celdaA25);
            LibroExcelModel.AplicarBordesARango(hojaDestino.Cells["A25:C31"]);

            int totalT1 = 0;
            int totalT2 = 0;
            int totalT3 = 0;
            int totalAltT1 = 0;
            int totalAltT3 = 0;

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    // Obtener el valor de la celda en la fila y columna actual
                    object cellValue = hojaOrigen.Cells[row, col].Value;
                    Console.Write(cellValue + "\t");
                    if (cellValue != null)
                    {


                        switch (cellValue.ToString())
                        {
                            case "Certificación Itinerario T1":
                                totalT1++;
                                break;
                            case "Certificación Itinerario  T2":
                                totalT2++;
                                break;
                            case "Certificación Itinerario  T3":
                                totalT3++;
                                break;
                            case "Certificación Itinerario en Altura T1":
                                totalAltT1++;
                                break;
                            case "Certificación Itinerario en Altura T3":
                                totalAltT3++;
                                break;

                        }


                    }
                    // MessageBox.Show(cellValue + "\t");
                }
                // Console.WriteLine(); // Nueva línea después de cada fila
            }

            hojaDestino.Cells["B26"].Value = totalT1;
            hojaDestino.Cells["B27"].Value = totalT2;
            hojaDestino.Cells["B28"].Value = totalT3;
            hojaDestino.Cells["B29"].Value = totalAltT1;
            hojaDestino.Cells["B30"].Value = totalAltT3;
            hojaDestino.Cells["B31"].Value = totalT1 + totalT2 + totalT3 + totalAltT1 + totalAltT3;

            double importeT1 = totalT1 * baremos.T1 * 2;
            double importeT2 = totalT2 * baremos.T2 * 2;
            double importeT3 = totalT3 * baremos.T3 * 2;
            double importeAltT1 = totalAltT1 * baremos.AlturaT1 * 2;
            double importeAltT3 = totalAltT3 * baremos.AlturaT3 * 2;

            hojaDestino.Cells["C26"].Value = $"$ {importeT1}";
            hojaDestino.Cells["C27"].Value = $"$ {importeT2}";
            hojaDestino.Cells["C28"].Value = $"$ {importeT3}";
            hojaDestino.Cells["C29"].Value = $"$ {importeAltT1}";
            hojaDestino.Cells["C30"].Value = $"$ {importeAltT3}";
            hojaDestino.Cells["C31"].Value = "$ " + (importeT1 + importeT2 + importeT3 + importeAltT1 + importeAltT3);
        }

        public void CrearTablaTotales(ExcelWorksheet hoja)
        {
            hoja.Cells["A35"].Value = "Descripción";
            hoja.Cells["A36"].Value = "Anomalias de Facturacion NC";
            hoja.Cells["A37"].Value = "Reclamos procedentes T1";
            hoja.Cells["A38"].Value = "Reclamos procedentes T2";
            hoja.Cells["A39"].Value = "Total de NC por Metodo Lineal (0,15% al 0,3%)";
            hoja.Cells["A40"].Value = "Totales Certificado";

            hoja.Cells["B35"].Value = "TOTAL";
            hoja.Cells["B36"].Value = 0;
            hoja.Cells["B37"].Value = 0;
            hoja.Cells["B38"].Value = 0;
            hoja.Cells["B39"].Value = 0;
            hoja.Cells["B40"].Value = 0;

            hoja.Cells["C35"].Value = "IMPORTE";
            hoja.Cells["C36"].Value = 0;
            hoja.Cells["C37"].Value = 0;
            hoja.Cells["C38"].Value = 0;
            hoja.Cells["C39"].Value = 0;
            hoja.Cells["C40"].Value = 0;

            hoja.Cells["D40"].Value = 0;

            LibroExcelModel.AplicarBordesARango(hoja.Cells["A35:C40"]);
        }

        public void CrearTablaValorFinalMulta(ExcelWorksheet hoja)
        {
            hoja.Cells["A44"].RichText.Add("Multa").Bold = true;
            hoja.Cells["B44"].Value = 0;
            hoja.Cells["C44"].Value = 0;

            hoja.Cells["A44:B44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            hoja.Cells["A44:B44"].Style.Fill.BackgroundColor.SetColor(Color.Orange);
        }
    }
}
