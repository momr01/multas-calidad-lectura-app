using Aspose.Cells;
using MultasLectura.Enums;
using MultasLectura.Helpers;
using MultasLectura.Models;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Services
{
    public class CalidadHojaResLecturistaService
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
            return new() {
                new ColorModel("rojo", Color.FromArgb(1, 255, 199, 206), Color.FromArgb(1, 156, 0, 6)),
                 new ColorModel("verde", Color.FromArgb(1, 198, 239, 206), Color.FromArgb(1, 0, 97, 0)),
                 new ColorModel("amarillo", Color.FromArgb(1, 255, 235, 156),Color.FromArgb(1, 156, 101, 0)),

            };
        }

        public void ColumnaLecturistaA(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"A{numPrimeraCelda}", empleado.Nombre, TipoOpCelda.Value);
        }

        public void ColumnaLeidosB(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"B{numPrimeraCelda}", empleado.Leidos, TipoOpCelda.Value);
        }

        public void ColumnaInconformidadesC(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"C{numPrimeraCelda}", empleado.Inconformidades, TipoOpCelda.Value);
        }

        public void CalcularTotal(ExcelWorksheet hoja, char letraCelda, int numCelda, int valor)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letraCelda.ToString().ToUpper()}{numCelda}", valor, TipoOpCelda.Value);
        }

        public void ColumnaIncXOpD(ExcelWorksheet hoja, int numPrimeraCelda)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"D{numPrimeraCelda}", $"C{numPrimeraCelda}/B{numPrimeraCelda}", TipoOpCelda.Formula);
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"D{numPrimeraCelda}"]);
        }

        public void ColumnaIncXNcE(ExcelWorksheet hoja, int numPrimeraCelda, int totalInconformidades)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"E{numPrimeraCelda}", $"C{numPrimeraCelda}/{totalInconformidades}", TipoOpCelda.Formula);
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"E{numPrimeraCelda}"]);
        }

        public void ColumnaAcumuladoF(int i, ExcelWorksheet hoja, int numPrimeraCelda)
        {
            if (i == 0)
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numPrimeraCelda}", $"+E{numPrimeraCelda}", TipoOpCelda.Formula);
            }
            else
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numPrimeraCelda}", $"+E{numPrimeraCelda}+F{numPrimeraCelda - 1}", TipoOpCelda.Formula);
            }
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"F{numPrimeraCelda}"]);
        }

        public double ColumnaIdealG(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado)
        {
            double idealPorcentaje = 0.0015;
            double ideal = empleado.Leidos * idealPorcentaje;
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numPrimeraCelda}", $"{ideal}", TipoOpCelda.Value);

            if (double.TryParse(hoja.Cells[$"G{numPrimeraCelda}"].Value?.ToString(), out double valor))
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numPrimeraCelda}", (int)Math.Round(valor), TipoOpCelda.Value);
            }
            return ideal;
        }

        public void ColumnaIncXOpIdealH(ExcelWorksheet hoja, int numPrimeraCelda)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"H{numPrimeraCelda}", "0,0015", TipoOpCelda.Value);

            if (double.TryParse(hoja.Cells[$"H{numPrimeraCelda}"].Value?.ToString(), out double valor2))
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"H{numPrimeraCelda}", valor2, TipoOpCelda.Value);
            }

            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"H{numPrimeraCelda}"]);
        }

        public double ColumnaDesvioI(ExcelWorksheet hoja, int numPrimeraCelda, EmpleadoModel empleado, double ideal, double totalIdeal)
        {
            double desvio = (ideal - empleado.Inconformidades) / totalIdeal;

            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"I{numPrimeraCelda}", desvio, TipoOpCelda.Value);
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"I{numPrimeraCelda}"]);

            return desvio;
        }

        public void ColorearSegunDesvio(ExcelWorksheet hoja, int numPrimeraCelda, List<ColorModel> colores, double desvio)
        {
            if (Math.Round(desvio, 4) <= -0.045)
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("rojo")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("rojo")).FirstOrDefault()!);
            }
            else if (Math.Round(desvio, 4) >= -0.0449 && Math.Round(desvio, 4) <= -0.001)
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("amarillo")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("amarillo")).FirstOrDefault()!);
            }
            else
            {
                LibroExcelHelper.ColorFondoLetra(hoja, 'i', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("verde")).FirstOrDefault()!);
                LibroExcelHelper.ColorFondoLetra(hoja, 'f', numPrimeraCelda, colores.Where(color => color.Nombre.Contains("verde")).FirstOrDefault()!);
            }
        }
    }
}
