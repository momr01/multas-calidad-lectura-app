using MultasLectura.Controlador.Interfaz;
using MultasLectura.Modelo;
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

namespace MultasLectura.Controlador
{
    public class Empleado
    {
         private string nombre;
         private int leidos;
         private int inconformidades;
         private double proporcion;
        //private double proporcion;

        public string Nombre { get { return nombre; } set { nombre = value; } }
        public int Leidos { get { return leidos; } set { leidos = value; } }
        public int Inconformidades { get { return inconformidades; } set { inconformidades = value; } }
        public double Proporcion { get { return proporcion; } set { proporcion = value; } }

        public void CalcularProporcion()
        {
            proporcion = (double)inconformidades / leidos;
        }

        public Empleado(string Nombre, int Leidos, int Inconformidades) { 
            this.Nombre = Nombre;
            this.Leidos = Leidos;
            this.Inconformidades = Inconformidades;
        }

      


      
       
     /*   public int CantidadT1 { get { return cantidadT1; } set { cantidadT1 = value; } }
        public int CantidadT2 { get { return cantidadT2; } set { cantidadT2 = value; } }
        public int CantidadT3 { get { return cantidadT3; } set { cantidadT3 = value; } }
        public int CantidadAlturaT1 { get { return cantidadAlturaT1; } set { cantidadAlturaT1 = value; } }
        public int CantidadAlturaT3 { get { return cantidadAlturaT3; } set { cantidadAlturaT3 = value; } }
        public double ImporteT1 { get { return importeT1; } set { importeT1 = value; } }

        public void CalcularImporteT1(double baremo)
        {
            importeT1 = 2 * cantidadT1 * baremo;
        }*/
    }


    public class CalidadHojaResLecturistaController : ICalidadHojaResLecturistaController
    {
        public void CrearTablaLecturistaInconformidades(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            // Obtener el número total de filas y columnas en la hoja de cálculo
            int rowCount = hojaCantXOper.Dimension.Rows;
            int colCount = hojaCantXOper.Dimension.Columns;

            hojaDestino.Cells["A1"].Value = "Lecturista";
            hojaDestino.Cells["B1"].Value = "Leídos";
            hojaDestino.Cells["C1"].Value = "Inconformidades";
            hojaDestino.Cells["D1"].Value = "% inc x op";
            hojaDestino.Cells["E1"].Value = "% inc x nc";
            hojaDestino.Cells["F1"].Value = "Acumulado";
            hojaDestino.Cells["G1"].Value = "Ideal";
            hojaDestino.Cells["H1"].Value = "% inc x op";
            hojaDestino.Cells["I1"].Value = "Desvío";



            // ApplyBorders(celdaA25);
            //LibroExcelModel.AplicarBordesARango(hojaDestino.Cells["A25:C31"]);

            /* int totalT1 = 0;
             int totalT2 = 0;
             int totalT3 = 0;
             int totalAltT1 = 0;
             int totalAltT3 = 0;*/

            List<Empleado> empleados = new List<Empleado>();
            int numPrimeraCelda = 2;


            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    object cellValue = hojaCantXOper.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.Contains("SYMESA"))
                        {
                            /*if (!empleados.Nombre.Contains(cellValue.ToString()!))
                            {
                                empleados.Add(cellValue.ToString()!);
                            }*/
                            bool contieneTexto = empleados.Any(empleado => empleado.Nombre.Contains(cellValue.ToString()));

                            /*if (empleados.Contains(empleado => empleado.Nombre == cellValue.ToString()))
                            {

                            }*/
                            if (!contieneTexto)
                            {
                                empleados.Add(new Empleado(Nombre: cellValue.ToString(), Leidos: int.Parse(hojaCantXOper.Cells[row, 5].Value.ToString()), Inconformidades: 0));
                            } else
                            {
                                empleados.Where(empleado => empleado.Nombre.Contains(cellValue.ToString())).FirstOrDefault().Leidos += int.Parse(hojaCantXOper.Cells[row, 5].Value.ToString());
                            }
                          

                        }
                    }
                }
            }

            // Obtener el número total de filas y columnas en la hoja de cálculo
            int rowsTablaCalDetalles = hojaCalidadDetalles.Dimension.Rows;
            int colsTablaCalDetalles = hojaCalidadDetalles.Dimension.Columns;

            int totalInconformidades = 0;
            int totalLeidos = 0;
            double totalIdeal = 0;


            for (int row = 1; row <= rowsTablaCalDetalles; row++)
            {
                for (int col = 1; col <= colsTablaCalDetalles; col++)
                {
                    object cellValue = hojaCalidadDetalles.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.Contains("SYMESA"))
                        { 
                            totalInconformidades++;
                            /*if (!empleados.Nombre.Contains(cellValue.ToString()!))
                            {
                                empleados.Add(cellValue.ToString()!);
                            }*/
                            bool contieneTexto = empleados.Any(empleado => empleado.Nombre.Contains(cellValue.ToString()));

                            /*if (empleados.Contains(empleado => empleado.Nombre == cellValue.ToString()))
                            {

                            }*/
                            if (!contieneTexto)
                            {
                                empleados.Add(new Empleado(Nombre: cellValue.ToString(), Leidos: 0, Inconformidades: 1));
                            }
                            else
                            {
                                empleados.Where(empleado => empleado.Nombre.Contains(cellValue.ToString())).FirstOrDefault().Inconformidades += 1;
                            }


                        }
                    }
                }
            }


          
            /*

            foreach (Empleado empleado in empleados)
            {
               // double acumulado = 0;
                double incXOp = empleado.Inconformidades / empleado.Leidos;
               
                hojaDestino.Cells[$"A{numPrimeraCelda}"].Value = empleado.Nombre;
                hojaDestino.Cells[$"B{numPrimeraCelda}"].Value = empleado.Leidos;
                hojaDestino.Cells[$"C{numPrimeraCelda}"].Value = empleado.Inconformidades;
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
             
                 hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

                numPrimeraCelda++;
            }*/

            foreach(Empleado empleado in empleados)
            {
                empleado.CalcularProporcion();
                totalIdeal += empleado.Leidos * 0.0015;
                totalLeidos += empleado.Leidos;
            }

            List<Empleado> empleadosOrdenados = empleados.OrderByDescending(emp => emp.Proporcion).ToList();

            double idealPorcentaje = 0.0015;

            

            Color verdeLetra = Color.FromArgb(1, 0, 97, 0);
            Color verdeFondo = Color.FromArgb(1, 198, 239, 206);
            Color rojoLetra;
            Color rojoFondo;
            Color amarilloLetra;
            Color amarilloFondo;


            for (int i = 0; i < empleadosOrdenados.Count; i++)
            {
               // MessageBox.Show(empleadosOrdenados[i].Proporcion.ToString());
                // double acumulado = 0;
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


                if(i == 0)
                {
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Formula = $"+E{numPrimeraCelda}";
                } else
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

                if (Math.Round(desvio,4) <= -0.045)
                {
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1,255,199,206));
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 0, 6));

                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                } else if(Math.Round(desvio, 4) >= -0.0449 && Math.Round(desvio, 4) <= -0.001)
                {
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 235, 156));
                    hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 101, 0));

                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                } else
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

            LibroExcelModel.AplicarBordesARango(rangoHojaResLecturista);

            hojaDestino.Cells.AutoFitColumns();

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
