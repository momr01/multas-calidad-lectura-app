using MultasLectura.Helpers;
using MultasLectura.Interfaces;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controllers
{
    public class CalidadHojaResLecturistaController : ICalidadHojaResLecturistaController
    {
        private void CrearEncabezados(ExcelWorksheet hoja)
        {
            Dictionary<string, string> headers = new()
            {
                ["A"] = "Lecturista",
                ["B"] = "Leídos",
                ["C"] = "Inconformidades",
                ["D"] = "% inc x op",
                ["E"] = "% inc x nc",
                ["F"] = "Acumulado",
                ["G"] = "Ideal",
                ["H"] = "% inc x op",
                ["I"] = "Desvío"
            };
            var claves = headers.Keys;

            for (int i = 0; i < headers.Count; i++)
            {
                hoja.Cells[$"{claves.ElementAt(i)}1"].Value = headers[claves.ElementAt(i)];
            }
        }

        private List<EmpleadoModel> CrearListaEmpleados(ExcelWorksheet hoja)
        {
            int contFilas = hoja.Dimension.Rows;
          
            List<EmpleadoModel> empleados = new();

            int colEmpleado = LibroExcelHelper.ObtenerNumeroColumna(hoja, "empleado");
            int colValores = LibroExcelHelper.ObtenerNumeroColumna(hoja, "compute_0005");

            if (colEmpleado != -1 && colValores != -1)
            {
                for (int fila = 1; fila <= contFilas; fila++)
                {
                        object cellValue = hoja.Cells[fila, colEmpleado].Value;
                        if (cellValue != null)
                        {
                            if (cellValue.ToString()!.ToLower().Contains("symesa"))
                            {
                                bool contieneTexto = empleados.Any(empleado => empleado.Nombre.Contains(cellValue.ToString()!));

                                if (!contieneTexto)
                                {
                                EmpleadoModel nuevoEmpleado = new(
                                    Nombre: cellValue.ToString()!,
                                    Leidos: int.Parse(hoja.Cells[fila, colValores].Value.ToString()!),
                                    Inconformidades: 0
                                );
                                
                                    empleados.Add(nuevoEmpleado);
                                }
                                else
                                {
                                EmpleadoModel emplExistente = empleados.Where(empleado => empleado.Nombre.Contains(cellValue.ToString()!)).FirstOrDefault()!;
                                emplExistente.Leidos += int.Parse(hoja.Cells[fila, colValores].Value.ToString()!);
                                }


                            }
                        }
                }
            }

            return empleados;
        }

        private void CalcularInconformidades(ExcelWorksheet hoja, 
          ref List<EmpleadoModel> empleados, 
          ref int totalInconformidades)
        {

            int contFilas = hoja.Dimension.Rows;
            int contColumnas = hoja.Dimension.Columns;
            int colEmpleado = LibroExcelHelper.ObtenerNumeroColumna(hoja, "lector");

            if (colEmpleado != -1)
            {
                for (int row = 1; row <= contFilas; row++)
                {
                    object cellValue = hoja.Cells[row, colEmpleado].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.ToLower().Contains("symesa"))
                        {
                            totalInconformidades++;
                            bool contieneTexto = empleados.Any(empleado => empleado.Nombre.Contains(cellValue.ToString()!));

                            if (!contieneTexto)
                            {
                                EmpleadoModel nuevoEmpleado = new(
                                    Nombre: cellValue.ToString()!,
                                    Leidos: 0,
                                    Inconformidades: 1
                                );

                                empleados.Add(nuevoEmpleado);
                            }
                            else
                            {
                                EmpleadoModel empleadoExistente = empleados.Where(empleado => empleado.Nombre.Contains(cellValue.ToString()!)).FirstOrDefault()!;
                                empleadoExistente.Inconformidades += 1;
                            }


                        }
                    }
                }
            }

        }

      private void CalcularProporcionIdealLeidos(ref List<EmpleadoModel> empleados, ref double totalIdeal, ref int totalLeidos)
        {

            foreach (EmpleadoModel empleado in empleados)
            {
                empleado.CalcularProporcion();
                totalIdeal += empleado.Leidos * 0.0015;
                totalLeidos += empleado.Leidos;
            }
        }

        private List<ColorModel> CargarColores()
        {
            /*Color verdeLetra = Color.FromArgb(1, 0, 97, 0);
            Color verdeFondo = Color.FromArgb(1, 198, 239, 206);
            Color rojoLetra = Color.FromArgb(1, 156, 0, 6);
            Color rojoFondo = Color.FromArgb(1, 255, 199, 206);
            Color amarilloLetra = Color.FromArgb(1, 156, 101, 0);
            Color amarilloFondo = Color.FromArgb(1, 255, 235, 156);*/

            return new() {
                new ColorModel("rojo", Color.FromArgb(1, 255, 199, 206), Color.FromArgb(1, 156, 0, 6)),
                 new ColorModel("verde", Color.FromArgb(1, 198, 239, 206), Color.FromArgb(1, 0, 97, 0)),
                 new ColorModel("amarillo", Color.FromArgb(1, 255, 235, 156),Color.FromArgb(1, 156, 101, 0)),

            };
        }

        public void CrearTablaLecturistaInconformidades(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            int numPrimeraCelda = 2;
            int totalInconformidades = 0;
            int totalLeidos = 0;
            double totalIdeal = 0;

            CrearEncabezados(hojaDestino);

            List<EmpleadoModel> empleados = CrearListaEmpleados(hojaCantXOper);

            CalcularInconformidades(hojaCalidadDetalles, ref empleados, ref totalInconformidades);

            CalcularProporcionIdealLeidos(ref empleados, ref totalIdeal, ref totalLeidos);
         
            List<EmpleadoModel> empleadosOrdenados = empleados.OrderByDescending(emp => emp.Proporcion).ToList();

            double idealPorcentaje = 0.0015;

            List<ColorModel> colores = CargarColores();


          /*  Color verdeLetra = Color.FromArgb(1, 0, 97, 0);
            Color verdeFondo = Color.FromArgb(1, 198, 239, 206);
            Color rojoLetra = Color.FromArgb(1, 156, 0, 6);
            Color rojoFondo = Color.FromArgb(1, 255, 199, 206);
            Color amarilloLetra = Color.FromArgb(1, 156, 101, 0);
            Color amarilloFondo = Color.FromArgb(1, 255, 235, 156);*/


            for (int i = 0; i < empleadosOrdenados.Count; i++)
            {
                double incXOp = empleadosOrdenados[i].Inconformidades / empleadosOrdenados[i].Leidos;


                hojaDestino.Cells[$"A{numPrimeraCelda}"].Value = empleadosOrdenados[i].Nombre;
                hojaDestino.Cells[$"B{numPrimeraCelda}"].Value = empleadosOrdenados[i].Leidos;
                hojaDestino.Cells[$"C{numPrimeraCelda}"].Value = empleadosOrdenados[i].Inconformidades;
                // hojaDestino.Cells[$"D{numPrimeraCelda}"].Value = empleado.Inconformidades / empleado.Leidos;
                // hojaDestino.Cells[$"D{numPrimeraCelda}"].Value = incXOp;
                hojaDestino.Cells[$"D{numPrimeraCelda}"].Formula = $"C{numPrimeraCelda}/B{numPrimeraCelda}";
                hojaDestino.Cells[$"D{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";
                // hojaDestino.Cells[$"E{numPrimeraCelda}"].Value = empleado.Inconformidades / totalInconformidades;
                hojaDestino.Cells[$"E{numPrimeraCelda}"].Formula = $"C{numPrimeraCelda}/{totalInconformidades}";
                hojaDestino.Cells[$"E{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

                // acumulado += empleado.Inconformidades / totalInconformidades;
                //MessageBox.Show(acumulado.ToString());
                // hojaDestino.Cells[$"F{numPrimeraCelda}"].Value = acumulado;
                // hojaDestino.Cells[$"F{numPrimeraCelda}"].Formula = $"+{acumulado}";


                if (i == 0)
                {
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Formula = $"+E{numPrimeraCelda}";
                }
                else
                {
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Formula = $"+E{numPrimeraCelda}+F{numPrimeraCelda - 1}";
                }

                hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";




                double ideal = empleadosOrdenados[i].Leidos * idealPorcentaje;

                //hojaDestino.Cells[$"G{numPrimeraCelda}"].Formula = $"+B{numPrimeraCelda}*{ideal}";
                hojaDestino.Cells[$"G{numPrimeraCelda}"].Value = $"{ideal}";

                if (double.TryParse(hojaDestino.Cells[$"G{numPrimeraCelda}"].Value?.ToString(), out double valor))
                {
                    // Asignar el valor convertido de vuelta a la celda
                    hojaDestino.Cells[$"G{numPrimeraCelda}"].Value = (int)Math.Round(valor);
                }



                hojaDestino.Cells[$"H{numPrimeraCelda}"].Value = "0,0015";

                if (double.TryParse(hojaDestino.Cells[$"H{numPrimeraCelda}"].Value?.ToString(), out double valor2))
                {
                    // Asignar el valor convertido de vuelta a la celda
                    hojaDestino.Cells[$"H{numPrimeraCelda}"].Value = valor2;
                }

                hojaDestino.Cells[$"H{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";


                double desvio = (ideal - empleadosOrdenados[i].Inconformidades) / 403.578;

                hojaDestino.Cells[$"I{numPrimeraCelda}"].Value = desvio;
                hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

                if (Math.Round(desvio, 4) <= -0.045)
                {
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 199, 206));
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 0, 6));

                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                }
                else if (Math.Round(desvio, 4) >= -0.0449 && Math.Round(desvio, 4) <= -0.001)
                {
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 235, 156));
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 101, 0));

                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                }
                else
                {
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 198, 239, 206));
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 0, 97, 0));

                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                }









                numPrimeraCelda++;

            }




            // Obtener el rango que cubre toda la tabla (incluyendo encabezados)
            /*  ExcelTable table = hojaDestino.Tables.FirstOrDefault();
              if (table == null)
              {
                  Console.WriteLine("No se encontró ninguna tabla en la hoja especificada.");
                  return;
              }


              // Definir el rango de la tabla (incluyendo encabezados)
              var start = hojaDestino.Dimension.Start;
              var end = hojaDestino.Dimension.End;
              ExcelRangeBase tableRange = hojaDestino.Cells[start.Row, start.Column, end.Row, end.Column];

              // Ordenar la tabla por una columna específica (por ejemplo, Columna2 en orden ascendente)
              // Obtener la columna que se utilizará para ordenar
              int columnIndex = 4; // Columna 2 (Columna2)

              // Aplicar el orden ascendente (true) a la columna especificada
              tableRange = (ExcelRangeBase)tableRange.OrderBy(cell => cell.Start.Row == start.Row ? null : hojaDestino.Cells[cell.Start.Row, columnIndex]);*/

            // Guardar los cambios en el archivo Excel






            // Obtener el valor de la celda en la fila y columna actual
            /* object cellValue = hojaOrigen.Cells[row, col].Value;
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


             }*/
            // MessageBox.Show(cellValue + "\t");
            // }
            // Console.WriteLine(); // Nueva línea después de cada fila
            //  }

            hojaDestino.Cells[$"B{empleados.Count + 2}"].Value = totalLeidos;
            hojaDestino.Cells[$"C{empleados.Count + 2}"].Value = totalInconformidades;
            hojaDestino.Cells[$"G{empleados.Count + 2}"].Value = (int)Math.Round(totalIdeal);


            MessageBox.Show(empleados.Count.ToString());

            var rangoHojaResLecturista = hojaDestino.Cells[hojaDestino.Dimension.Address];

            LibroExcelHelper.AplicarBordeFinoARango(rangoHojaResLecturista);

            //hojaDestino.Cells.AutoFitColumns();

            /*  hojaDestino.Cells["B26"].Value = totalT1;
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
              hojaDestino.Cells["C31"].Value = "$ " + (importeT1 + importeT2 + importeT3 + importeAltT1 + importeAltT3);*/
        }
    }
}
