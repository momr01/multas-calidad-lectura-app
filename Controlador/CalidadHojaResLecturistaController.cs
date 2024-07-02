using MultasLectura.Controlador.Interfaz;
using MultasLectura.Modelo;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controlador
{
    public class Empleado
    {
        public string Nombre { get; set; }
        public int Leidos { get; set; }
        public int Inconformidades { get; set; }

        public Empleado(string Nombre, int Leidos, int Inconformidades) { 
            this.Nombre = Nombre;
            this.Leidos = Leidos;
            this.Inconformidades = Inconformidades;
        }
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

            foreach (Empleado empleado in empleados)
            {
                hojaDestino.Cells[$"A{numPrimeraCelda}"].Value = empleado.Nombre;
                hojaDestino.Cells[$"B{numPrimeraCelda}"].Value = empleado.Leidos;
                hojaDestino.Cells[$"C{numPrimeraCelda}"].Value = empleado.Inconformidades;
                numPrimeraCelda++;
            }

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

            MessageBox.Show(empleados.Count.ToString());

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
