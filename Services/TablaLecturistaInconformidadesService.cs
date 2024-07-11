using MultasLectura.Helpers;
using MultasLectura.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Services
{
    public class TablaLecturistaInconformidadesService
    {
        public void CrearEncabezados(ExcelWorksheet hoja)
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

        public List<EmpleadoModel> CrearListaEmpleados(ExcelWorksheet hoja)
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

        public void CalcularInconformidades(ExcelWorksheet hoja,
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

        public void CalcularProporcionIdealLeidos(ref List<EmpleadoModel> empleados, ref double totalIdeal, ref int totalLeidos)
        {

            foreach (EmpleadoModel empleado in empleados)
            {
                empleado.CalcularProporcion();
                totalIdeal += empleado.Leidos * 0.0015;
                totalLeidos += empleado.Leidos;
            }
        }

        public List<ColorModel> CargarColores()
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

        public void ColumnaLecturistaA(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            hoja.Cells[$"A{numPrimeraCelda}"].Value = empleado.Nombre;

        }

        public void ColumnaLeidosB(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            hoja.Cells[$"B{numPrimeraCelda}"].Value = empleado.Leidos;

        }

        public void ColumnaInconformidadesC(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            hoja.Cells[$"C{numPrimeraCelda}"].Value = empleado.Inconformidades;

        }

        public void CalcularTotal(ExcelWorksheet hoja, char letraCelda, int numCelda, int valor)
        {
            hoja.Cells[$"{letraCelda.ToString().ToUpper()}{numCelda}"].Value = valor;
            // hojaDestino.Cells[$"C{empleados.Count + 2}"].Value = totalInconformidades;
            //  hojaDestino.Cells[$"G{empleados.Count + 2}"].Value = (int)Math.Round(totalIdeal);
        }

        public void ColumnaIncXOpD(ExcelWorksheet hoja, int numPrimeraCelda)
        {
            hoja.Cells[$"D{numPrimeraCelda}"].Formula = $"C{numPrimeraCelda}/B{numPrimeraCelda}";
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"D{numPrimeraCelda}"]);
            //  hoja.Cells[$"D{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

        }

        public void ColumnaIncXNcE(ExcelWorksheet hoja, int numPrimeraCelda, int totalInconformidades)
        {
            hoja.Cells[$"E{numPrimeraCelda}"].Formula = $"C{numPrimeraCelda}/{totalInconformidades}";
            hoja.Cells[$"E{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

        }

        public void ColumnaAcumuladoF(int i, ExcelWorksheet hoja, int numPrimeraCelda)
        {
            if (i == 0)
            {
                hoja.Cells[$"F{numPrimeraCelda}"].Formula = $"+E{numPrimeraCelda}";
            }
            else
            {
                hoja.Cells[$"F{numPrimeraCelda}"].Formula = $"+E{numPrimeraCelda}+F{numPrimeraCelda - 1}";
            }

            hoja.Cells[$"F{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

        }

        public double ColumnaIdealG(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            double idealPorcentaje = 0.0015;
            double ideal = empleado.Leidos * idealPorcentaje;

            //hojaDestino.Cells[$"G{numPrimeraCelda}"].Formula = $"+B{numPrimeraCelda}*{ideal}";
            hoja.Cells[$"G{numPrimeraCelda}"].Value = $"{ideal}";

            if (double.TryParse(hoja.Cells[$"G{numPrimeraCelda}"].Value?.ToString(), out double valor))
            {
                // Asignar el valor convertido de vuelta a la celda
                hoja.Cells[$"G{numPrimeraCelda}"].Value = (int)Math.Round(valor);
            }

            return ideal;

        }

        public void ColumnaIncXOpIdealH(ExcelWorksheet hoja, int numPrimeraCelda)
        {
            hoja.Cells[$"H{numPrimeraCelda}"].Value = "0,0015";

            if (double.TryParse(hoja.Cells[$"H{numPrimeraCelda}"].Value?.ToString(), out double valor2))
            {
                // Asignar el valor convertido de vuelta a la celda
                hoja.Cells[$"H{numPrimeraCelda}"].Value = valor2;
            }

            hoja.Cells[$"H{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

        }

        public double ColumnaDesvioI(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado, double ideal, double totalIdeal)
        {
            // double desvio = (ideal - empleado.Inconformidades) / 403.578;
            double desvio = (ideal - empleado.Inconformidades) / totalIdeal;

            hoja.Cells[$"I{numPrimeraCelda}"].Value = desvio;
            hoja.Cells[$"I{numPrimeraCelda}"].Style.Numberformat.Format = "0.00%";

            return desvio;

        }

        public void ColorearSegunDesvio(ExcelWorksheet hoja, int numPrimeraCelda, List<ColorModel> colores, double desvio)
        {
            if (Math.Round(desvio, 4) <= -0.045)
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("rojo")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("rojo")).FirstOrDefault()!);

                /* hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                 hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 199, 206));
                 hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 0, 6));

                 hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                 hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);*/
            }
            else if (Math.Round(desvio, 4) >= -0.0449 && Math.Round(desvio, 4) <= -0.001)
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("amarillo")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("amarillo")).FirstOrDefault()!);

                /*hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 255, 235, 156));
                hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 156, 101, 0));

                hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);*/
            }
            else
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("verde")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("verde")).FirstOrDefault()!);
                /* hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                 hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 198, 239, 206));
                 hojaDestino.Cells[$"I{numPrimeraCelda}"].Style.Font.Color.SetColor(Color.FromArgb(1, 0, 97, 0));

                 hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                 hojaDestino.Cells[$"F{numPrimeraCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);*/
            }

        }

    }
}
