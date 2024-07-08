using Aspose.Cells;
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

namespace MultasLectura.Helpers
{
    public class LibroExcelHelper
    {
        //static public void IniciarProcesoCarga(System.Windows.Forms.TextBox txt, System.Action<string> funcionCargarLibro)
        static public void IniciarProcesoCarga(TextBox txt)
        {
            string filePath = CargarLibroExcel();

            if (string.IsNullOrEmpty(filePath))
            {
                txt.Text = string.Empty;
                MostrarMensaje("Ocurrió un error al intentar cargar el archivo. Por favor inténtelo nuevamente", true);
            }
            else
            {
                txt.Text = filePath;
                //funcionCargarLibro(filePath);
            }
        }

        static public void AplicarBordeFinoARango(ExcelRangeBase rango)
        {
            rango.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        static public void AplicarBordeGruesoARango(ExcelRangeBase rango)
        {
            rango.Style.Border.Top.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Left.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Right.Style = ExcelBorderStyle.Thick;
        }

        static public void FormatoMoneda(ExcelRange rango)
        {
            using (ExcelRange celda = rango)
            {
                celda.Style.Numberformat.Format = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
            }
        }

        static public void FormatoPorcentaje(ExcelRange rango)
        {
            rango.Style.Numberformat.Format = "0.00%";
        }

        static public void FormatoNegrita(ExcelRange rango)
        {
            rango.Style.Font.Bold = true;
        }

        static public void FondoSolido(ExcelRange rango, Color color)
        {
            rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rango.Style.Fill.BackgroundColor.SetColor(color);
        }

        static public int ObtenerNumeroColumna(ExcelWorksheet hoja, string encabezado)
        {
            int contadorColumnas = hoja.Dimension.End.Column;
            for (int col = 1; col <= contadorColumnas; col++)
            {
                if (hoja.Cells[1, col].Text == encabezado)
                {
                    return col;
                }
            }
            return -1;
        }


        static public string ValidarFormato(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string fileExtension = Path.GetExtension(filePath);

            if (fileExtension.ToLower() == ".xlsx")
            {
                return filePath;
            }
            else if (fileExtension.ToLower() == ".xls")
            {
                
                string convertedFilePath = ConvertirXlsAXlsx(filePath);
                if (!string.IsNullOrEmpty(convertedFilePath))
                {
                    return convertedFilePath;
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        static private string ConvertirXlsAXlsx(string xlsFilePath)
        {
           
            try
            {

                // Cargar archivo .xls
                using (FileStream fs = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfwb = new HSSFWorkbook(fs); // Crear instancia de libro .xls
                    XSSFWorkbook workbook = new XSSFWorkbook(); // Crear instancia de libro .xlsx


                    // Copiar hojas de .xls a .xlsx
                    for (int i = 0; i < hssfwb.NumberOfSheets; i++)
                    {
                        ISheet sheet = hssfwb.GetSheetAt(i);
                        XSSFSheet newSheet = (XSSFSheet)workbook.CreateSheet(sheet.SheetName);

                        // Copiar filas y celdas
                        for (int j = 0; j <= sheet.LastRowNum; j++)
                        {
                            IRow row = sheet.GetRow(j);
                            XSSFRow newRow = (XSSFRow)newSheet.CreateRow(j);

                            if (row != null)
                            {
                                for (int k = 0; k < row.LastCellNum; k++)
                                {
                                    ICell cell = row.GetCell(k);
                                    if (cell != null)
                                    {
                                        XSSFCell newCell = (XSSFCell)newRow.CreateCell(k);
                                        newCell.SetCellValue(cell.ToString());
                                    }
                                }
                            }
                        }
                    }

                    var partesPath = xlsFilePath.Split('\\');
                    string pathGuardarConversion = "";

                    foreach (string parte in partesPath)
                    {
                        if (partesPath[partesPath.Length - 1] != parte)
                        {
                            pathGuardarConversion += parte + '\\';
                        }
                    }

                    string nombreOriginal = Path.GetFileNameWithoutExtension(xlsFilePath);

                    //   Random random = new Random();

                    // Generar un número aleatorio entre 1 y 100 (ambos inclusive)
                    //  int numeroAleatorio = random.Next(1, 101);

                    //  pathGuardarConversion += $"libro-conversion{numeroAleatorio}.xlsx";
                   // DateTime.Now.ToString("yyyyMMdd") + "-" + DateTime.Now.ToString("HHmmss")
                    pathGuardarConversion += $"{nombreOriginal}_{DateTime.Now.ToString("yyyyMMdd")}{DateTime.Now.ToString("HHmmss")}.xlsx";

                    //  MessageBox.Show(partesPath[partesPath.Length - 1]);

                    // Guardar como .xlsx
                    using (FileStream fileOut = new FileStream(pathGuardarConversion, FileMode.Create))
                    {
                        workbook.Write(fileOut);
                    }

                    return pathGuardarConversion;
                }


            }
            catch
            {
                return string.Empty;
            }
        }

        static public string CargarLibroExcel()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                // Configurar propiedades del OpenFileDialog
                openFileDialog.InitialDirectory = "c:\\"; // Directorio inicial
                //openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx|Archivos Excel (*.xls)|*.xls|Todos los archivos (*.*)|*.*"; // Filtros de archivo
                openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx|Archivos Excel (*.xls)|*.xls";
                openFileDialog.FilterIndex = 1; // Índice del filtro predeterminado
                openFileDialog.RestoreDirectory = true; // Restaurar el directorio anterior al cerrar el diálogo

                // Mostrar el diálogo y verificar si el usuario ha seleccionado un archivo
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Mostrar la ruta del archivo seleccionado en el TextBox
                    // txtRutaCalidadDetalles.Text = openFileDialog.FileName;

                    // AbrirArchivo2(openFileDialog.FileName);
                    //   calidadController.CargarLibroCalidadDetalles(openFileDialog.FileName);

                    // Aquí puedes realizar cualquier operación adicional con el archivo seleccionado
                    // Por ejemplo, cargar y procesar el archivo Excel usando EPPlus como se mostró anteriormente
                    return openFileDialog.FileName;
                }
                else
                {
                    //txtRutaCalidadDetalles.Text = string.Empty;
                    return "";
                }
            }
            catch
            {
                return "";
            }

        }

        static public void ConvertirTextoANumero(ExcelRange rango)
        {
            foreach (var celda in rango)
            {
                if (double.TryParse(celda.Value?.ToString(), out double valor))
                {
                    // Asignar el valor convertido de vuelta a la celda
                    celda.Value = valor;
                }
            }
        }

        static public string DialogoGuardarArchivo()
        {
            string rutaArchivo = string.Empty;

            SaveFileDialog guardar = new SaveFileDialog();
            guardar.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            guardar.FilterIndex = 1;
            guardar.RestoreDirectory = true;

            guardar.FileName = "+ Calidad_Lectura.xlsx";

            if (guardar.ShowDialog() == DialogResult.OK)
            {
                rutaArchivo = guardar.FileName;
            }

            return rutaArchivo;
        }

        static public void MostrarMensaje(string mensaje, bool esError)
        {
            MessageBox.Show(mensaje, esError ? "ERROR" : "ATENCIÓN", MessageBoxButtons.OK, esError ? MessageBoxIcon.Error : MessageBoxIcon.Warning);
        }

        static public int SumarColumnaInt(ExcelWorksheet hoja, int colASumar, int filaInicial)
        {
            int totalSuma = 0;

            // int columnToSum = 5; // Columna que quieres sumar (por ejemplo, A=1, B=2, etc.)
            //int startRow = 2; // Fila inicial donde comienzan los datos (puede variar según el archivo)
            int filaFinal = hoja.Dimension.End.Row; // Última fila con datos en la hoja

            // double total = 0;

            for (int fila = filaInicial; fila <= filaFinal; fila++)
            {
                var celda = hoja.Cells[fila, colASumar].Value;
                if (celda != null)
                {
                    int datoCelda;
                    if (int.TryParse(celda.ToString(), out datoCelda))
                    {
                        totalSuma += datoCelda;
                    }
                    // Si los datos no son numéricos, se pueden ignorar o manejar según sea necesario
                }
            }

            return totalSuma;
        }


    }
}
