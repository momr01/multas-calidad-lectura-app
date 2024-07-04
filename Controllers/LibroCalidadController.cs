using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using MultasLectura.Helpers;
using MultasLectura.Interfaces;
using MultasLectura.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controllers
{
    public class LibroCalidadController : ILibroCalidadController
    {
        private readonly BaremoModel _baremos;
        private readonly MetaModel _metas;
        // private Multa multa;
        private readonly ICalidadHojaResumenController _hojaResumenController;
        private readonly ICalidadHojaCuadrosController _hojaCuadrosController;
        private readonly ICalidadHojaResLecturistaController _hojaResLecturistaController;



        public LibroCalidadController(BaremoModel baremos, MetaModel metas)
        {
            _hojaResumenController = new CalidadHojaResumenController();
            _hojaCuadrosController = new CalidadHojaCuadrosController();
            _hojaResLecturistaController = new CalidadHojaResLecturistaController();
            _baremos = baremos;
            _metas = metas;
            /*baremos = new BaremoModel
            {
                T1 = 304.91,
                T2 = 2814.51,
                T3 = 304.91,
                AlturaT1 = 3572.63,
                AlturaT3 = 3572.63
            };*/
            // multa = new Multa();
        }

        public void CargarLibroExcel(string pathCalidadDetalles, string pathCalXOper, double importeCertificacion)
        {
            string archivoAUtilizar = LibroExcelHelper.ValidarFormato(pathCalidadDetalles);

            string archivoCalXOper = LibroExcelHelper.ValidarFormato(pathCalXOper);

            if (string.IsNullOrEmpty(archivoAUtilizar) || string.IsNullOrEmpty(archivoCalXOper))
            {
                LibroExcelHelper.MostrarMensaje("Error al convertir el archivo .xls a .xlsx", true);
            }
            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                // Configurar el diálogo de guardar archivo
                saveFileDialog1.Filter = "Archivos Excel (*.xlsx)|*.xlsx";// Filtro de tipos de archivo
                saveFileDialog1.FilterIndex = 1; // Índice del filtro predeterminado
                saveFileDialog1.RestoreDirectory = true; // Restaurar el directorio al cerrar el diálogo

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Obtener la ruta del archivo seleccionado
                    string rutaArchivo = saveFileDialog1.FileName;

                    // Aquí puedes realizar acciones adicionales, como guardar el archivo
                    // Por ejemplo, guardar un archivo de texto
                    // File.WriteAllText(rutaArchivo, contenido);

                    // Mostrar un mensaje con la ruta seleccionada (opcional)
                    GenerarLibroCalidad(archivoAUtilizar, archivoCalXOper, importeCertificacion, rutaArchivo);
                    MessageBox.Show($"Archivo guardado en: {rutaArchivo}", "Guardar archivo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }








              //  GenerarLibroCalidad(archivoAUtilizar, archivoCalXOper, importeCertificacion);
            }

        }

        private void GenerarLibroCalidad(string filePath, string pathCalXOper, double importeCertificacion, string rutaGuardar)
        {
            // Abrir y leer el archivo Excel con EPPlus
            using (ExcelPackage excelPackageCalXOper = new ExcelPackage(new FileInfo(pathCalXOper)))
            {
                ExcelWorksheet hojaBaseCalXOp = excelPackageCalXOper.Workbook.Worksheets[0];
                // Abrir y leer el archivo Excel con EPPlus
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Obtener la primera hoja de cálculo del archivo
                    ExcelWorksheet hojaBase = excelPackage.Workbook.Worksheets[0];

                    // Obtener el número total de filas y columnas en la hoja de cálculo
                    int rowCount = hojaBase.Dimension.Rows;
                    int colCount = hojaBase.Dimension.Columns;

                    //crear rango para analizar
                    //  var rangoCalidadDetalles = hojaBase.Cells[hojaBase.Dimension.Address];

                    //creamos hojas nuevas del libro
                    ExcelWorksheet hojaResumen = excelPackage.Workbook.Worksheets.Add("Resumen");
                    ExcelWorksheet hojaResLecturista = excelPackage.Workbook.Worksheets.Add("Res-Lecturista");
                    ExcelWorksheet hojaCantXOperario = excelPackage.Workbook.Worksheets.Add("Cant_x_Oper", hojaBaseCalXOp);
                    ExcelWorksheet hojaCuadros = excelPackage.Workbook.Worksheets.Add("Cuadros");
                    ExcelWorksheet hojaEliminados = excelPackage.Workbook.Worksheets.Add("ELIMINADOS");


                  

                    // libro2Package.Workbook.Worksheets.Add(sheet.Name, sheet);

                    //ubicacion de hojas
                    excelPackage.Workbook.Worksheets.MoveBefore("Resumen", "calidad_detalle");
                    excelPackage.Workbook.Worksheets.MoveBefore("Res-Lecturista", "calidad_detalle");


                    // Obtener el rango de celdas en la hoja copiada
                    var rangoHojaCantXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];
                    LibroExcelHelper.ConvertirTextoANumero(rangoHojaCantXOperario);


                    // Obtener el rango de celdas en la hoja copiada
                    /* var rangoCeldas = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];

                     // Convertir texto a número en el rango de celdas
                     //  rangoCeldas.TextToNumber();
                     foreach (var cell in rangoCeldas)
                     {
                         if (double.TryParse(cell.Value?.ToString(), out double valor))
                         {
                             // Asignar el valor convertido de vuelta a la celda
                             cell.Value = valor;
                         }
                     }*/




                    ////////////////
                    ///
                    /* var columnCells = hojaCantXOperario.Cells["E:E"];

                     // Iterar sobre las celdas de la columna y convertir los valores a números
                     foreach (var cell in columnCells)
                     {
                         // Convertir el valor de la celda a número
                         if (double.TryParse(cell.Value?.ToString(), out double valor))
                         {
                             // Hacer algo con el número convertido, por ejemplo, imprimirlo
                             Console.WriteLine($"Valor convertido en la celda {cell.Address}: {valor}");
                         }
                         else
                         {
                             // Manejo de errores si no se puede convertir el valor a número
                             Console.WriteLine($"El valor en la celda {cell.Address} no es un número válido.");
                         }
                     }*/







                    ////////////


                    //crear rango para analizar
                    var rangoCalidadDetalles = hojaBase.Cells[hojaBase.Dimension.Address];
                    var rangoCalXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];

                    // hojaCantXOperario = excelPackageCalXOper.Workbook.Worksheets[0];

                    AgregarContenidoHojaResumen(hojaBase, hojaResumen, rangoCalidadDetalles, _baremos, _metas, hojaBaseCalXOp, importeCertificacion);

                    AgregarContenidoHojaCuadros(hojaCuadros, rangoCalidadDetalles, rangoCalXOperario);
                    AgregarContenidoHojaResLecturista(hojaCantXOperario, hojaBase, hojaResLecturista);
                    /* AgregarContenidoHojaResLecturista();
                     AgregarContenidoHojaCantXOperario();
                     AgregarContenidoHojaCuadros();*/


                    /*  CrearTablaDinTipoEstado(hojaResumen, rangoTablaDinamica);

                      CrearTablaMetodoLineal(hojaResumen, hojaBase);

                      CrearTablaTotales(hojaResumen);

                      CrearTablaValorFinalMulta(hojaResumen);

                      CrearTablaBaremosMetas(hojaResumen);*/


                    //  multa.CantidadT1 = 345;
                    //   multa.CalcularImporteT1(baremos.T1);

                    //  label1.Text = multa.ImporteT1.ToString();


                    /*  var tabla = hojaResumen.Tables["TablaDinTipoEstado"];
                      //var pivotTable = hojaResumen.PivotTables["TablaDinTipoEstado"];
                      var pivotTable = hojaResumen.PivotTables.FirstOrDefault(pt => pt.Name == "TablaDinTipoEstado");

                      // var rangofinal = pivotTable.TableRange;

                      //var rangoTabla = tabla.Address;

                      /*int startRow = rangoTabla.Start.Row;
                      int startColumn = rangoTabla.Start.Column;
                      int endRow = rangoTabla.End.Row;
                      int endColumn = rangoTabla.End.Column;*/

                    /*  var startRow = pivotTable.RowFields;


                    /*  int startColumn = rangoTabla.Start.Column;
                      int endRow = rangoTabla.End.Row;
                      int endColumn = rangoTabla.End.Column;*/

                    //  MessageBox.Show(startRow + " - columna= " );





                    // Guardar archivo
                    //  excelPackage.SaveAs(new FileInfo(@"C:/Users/maxio/Documents/archivo3.xlsx"));
                    excelPackage.SaveAs(new FileInfo(rutaGuardar));


                }
            }

        }

        private void AgregarContenidoHojaResumen(
            ExcelWorksheet hojaBase, 
            ExcelWorksheet hojaResumen, 
            ExcelRange rango, 
            BaremoModel baremos, 
            MetaModel metas, 
            ExcelWorksheet hojaCalXOperario, 
            double importeCertificacion)
        {
            //  _hojaResumenController.Prueba();
            _hojaResumenController.CrearTablaDinTipoEstado(hojaResumen, rango);
            Dictionary<string, double> totales = _hojaResumenController.CrearTablaMetodoLineal(hojaResumen, hojaBase, baremos);
            Dictionary<string, double> propInMasImpMetLineal = _hojaResumenController.CrearTablaTotales(hojaResumen, totales, new() { ["t1"] = 43, ["t2"] = 5 }, baremos, hojaCalXOperario, importeCertificacion);
            _hojaResumenController.CrearTablaValorFinalMulta(hojaResumen, propInMasImpMetLineal["propInconformidades"], propInMasImpMetLineal["totalMetLineal"],  importeCertificacion, metas);
            _hojaResumenController.CrearTablaBaremosMetas(hojaResumen, baremos, metas, propInMasImpMetLineal["propInconformidades"]);
            hojaResumen.Cells.AutoFitColumns();
            hojaResumen.Column(2).AutoFit();

        }

        private void AgregarContenidoHojaCuadros(ExcelWorksheet hojaCuadros, ExcelRange rangoCalidadDetalles, ExcelRange rangoCalXOperario)
        {
            _hojaCuadrosController.CrearTablaDinEmpleadoTotal(hojaCuadros, rangoCalXOperario);
            _hojaCuadrosController.CrearTablaDinLectorTotal(hojaCuadros, rangoCalidadDetalles);
            hojaCuadros.Cells.AutoFitColumns();

        }

        private void AgregarContenidoHojaResLecturista(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            _hojaResLecturistaController.CrearTablaLecturistaInconformidades(hojaCantXOper, hojaCalidadDetalles, hojaDestino);
            hojaDestino.Cells.AutoFitColumns();
        }

        /*  private void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango)
          {
              // Crear tabla dinámica
              var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinTipoEstado");
              pivotTable.RowFields.Add(pivotTable.Fields["tipo_certificacion"]);
              pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
              pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
              pivotTable.DataFields[0].Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count;

              //label2.Text =  pivotTable.Fields.Count.ToString();



          }*/


        /*
        private void CrearTablaMetodoLineal(ExcelWorksheet hojaDestino, ExcelWorksheet hojaOrigen)
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

        }*/

        /*
        private void CrearTablaTotales(ExcelWorksheet hoja)
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
        */

        /*
        private void CrearTablaValorFinalMulta(ExcelWorksheet hoja)
        {
            hoja.Cells["A44"].RichText.Add("Multa").Bold = true;
            hoja.Cells["B44"].Value = 0;
            hoja.Cells["C44"].Value = 0;

            hoja.Cells["A44:B44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            hoja.Cells["A44:B44"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

        }
        */

        /*
        private void CrearTablaBaremosMetas(ExcelWorksheet hoja)
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
        */

        /*

         void ILibroCalidadController.CargarLibroCalidadDetalles(string filePath)
         {
             throw new NotImplementedException();
         }

         void ILibroCalidadController.CargarLibroReclamosDetalles(string filePath)
         {
             throw new NotImplementedException();
         }

         void ILibroCalidadController.CargarLibroCalidadXOperario(string filePath)
         {
             throw new NotImplementedException();
         }

         void ILibroCalidadController.CargarBaremos()
         {
             throw new NotImplementedException();
         }

         void ILibroCalidadController.CargarMetas()
         {
             throw new NotImplementedException();
         }
        */



    }
}
