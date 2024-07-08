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
        }

        public void CargarLibroExcel(string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion)
        {
            try
            {
                string archCalDetalles = LibroExcelHelper.ValidarFormato(rutaCalDetalles);
                string archCalXOperario = LibroExcelHelper.ValidarFormato(rutaCalXOper);
                string archReclDetalles = LibroExcelHelper.ValidarFormato(rutaReclDetalles);

                if (string.IsNullOrEmpty(archCalDetalles) || string.IsNullOrEmpty(archCalXOperario) 
                    || string.IsNullOrEmpty(archReclDetalles)
                    )
                {
                    LibroExcelHelper.MostrarMensaje("Error al cargar los archivos. Intente nuevamente.", true);
                }
                else
                {
                    string rutaArchivo = LibroExcelHelper.DialogoGuardarArchivo();

                    if (string.IsNullOrEmpty(rutaArchivo))
                    {
                        LibroExcelHelper.MostrarMensaje("Tarea cancelada por el usuario.", true);
                    }
                    else
                    {
                        GenerarLibroCalidad(archCalDetalles, archCalXOperario, archReclDetalles, importeCertificacion, rutaArchivo);
                    }

                }
            } catch (Exception e) {
                LibroExcelHelper.MostrarMensaje(e.Message, true);
            }
           

        }

   
        private Dictionary<string, int> ReclamosPorTarifa(ExcelWorksheet hoja, int numeroColumna)
        {
            int contFilas = hoja.Dimension.Rows;

            int totalReclT1 = 0;
            int totalReclT2 = 0;

            for (int row = 1; row <= contFilas; row++)
            {
                object cellValue = hoja.Cells[row, numeroColumna].Value;
                if (cellValue != null)
                {

                    if (cellValue.ToString().ToLower().Contains("t1"))
                    {
                        totalReclT1++;
                    }
                    else if (cellValue.ToString().ToLower().Contains("t2"))
                    {
                        totalReclT2++;
                    }

                }
            }

            return new()
            {
                ["t1"] = totalReclT1,
                ["t2"] = totalReclT2
            };
        } 



    private void GenerarLibroCalidad(string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion, string rutaGuardar)
        {
            using ExcelPackage libroCalXOperario = new(new FileInfo(rutaCalXOper));
            ExcelWorksheet hojaBaseCalXOp = libroCalXOperario.Workbook.Worksheets[0];

            using ExcelPackage libroReclDetalles = new(new FileInfo(rutaReclDetalles));
            ExcelWorksheet hojaBaseReclDetalles = libroReclDetalles.Workbook.Worksheets[0];

            using ExcelPackage libroCalDetalles = new(new FileInfo(rutaCalDetalles));
            ExcelWorksheet hojaBaseCalDetalles = libroCalDetalles.Workbook.Worksheets[0];


            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroCalDetalles.Workbook.Worksheets.Add("Resumen");
            ExcelWorksheet hojaResLecturista = libroCalDetalles.Workbook.Worksheets.Add("Res-Lecturista");
            ExcelWorksheet hojaCantXOperario = libroCalDetalles.Workbook.Worksheets.Add("Cant_x_Oper", hojaBaseCalXOp);
            ExcelWorksheet hojaCuadros = libroCalDetalles.Workbook.Worksheets.Add("Cuadros");
            ExcelWorksheet hojaEliminados = libroCalDetalles.Workbook.Worksheets.Add("ELIMINADOS");


            //ubicacion de hojas
            libroCalDetalles.Workbook.Worksheets.MoveBefore("Resumen", "calidad_detalle");
            libroCalDetalles.Workbook.Worksheets.MoveBefore("Res-Lecturista", "calidad_detalle");


            // Obtener el rango de celdas en la hoja copiada
            var rangoHojaCantXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];
            LibroExcelHelper.ConvertirTextoANumero(rangoHojaCantXOperario);


          //  int rowCount = hojaBaseReclDetalles.Dimension.Rows;
          //   int colCount = hojaBaseReclDetalles.Dimension.Columns;

          //  int totalReclT1 = 0;
          //  int totalReclT2 = 0;

            // Llama a la función para obtener el número de columna
            // int columnNumber = GetColumnNumberByHeader(filePath, headerName);
            int numeroColumna = LibroExcelHelper.ObtenerNumeroColumna(hojaBaseReclDetalles, "desc_tar");

            Dictionary<string, int> reclamosValores = new Dictionary<string, int>();

            if (numeroColumna != -1)
            {
                //Console.WriteLine($"El encabezado '{headerName}' se encuentra en la columna número {columnNumber}.");
                // MessageBox.Show("numero de columna: " + columnNumber);
                reclamosValores = ReclamosPorTarifa(hojaBaseReclDetalles, numeroColumna);
            }
            else
            {
                //Console.WriteLine($"El encabezado '{headerName}' no se encontró.");
                // MessageBox.Show("NO EXISTE LA COLUMNA");
                throw new Exception();
            }


            //crear rango para analizar
            var rangoCalidadDetalles = hojaBaseCalDetalles.Cells[hojaBaseCalDetalles.Dimension.Address];
            var rangoCalXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];

       

            AgregarContenidoHojaResumen(hojaBaseCalDetalles, hojaResumen, rangoCalidadDetalles, _baremos, _metas, hojaBaseCalXOp, importeCertificacion, reclamosValores);

            AgregarContenidoHojaCuadros(hojaCuadros, rangoCalidadDetalles, rangoCalXOperario);
            AgregarContenidoHojaResLecturista(hojaCantXOperario, hojaBaseCalDetalles, hojaResLecturista);
            
            libroCalDetalles.SaveAs(new FileInfo(rutaGuardar));

        }

        private void AgregarContenidoHojaResumen(
            ExcelWorksheet hojaBase, 
            ExcelWorksheet hojaResumen, 
            ExcelRange rango, 
            BaremoModel baremos, 
            MetaModel metas, 
            ExcelWorksheet hojaCalXOperario, 
            double importeCertificacion,
            Dictionary<string, int> reclamosValores
            )
        {
            _hojaResumenController.CrearTablaDinTipoEstado(hojaResumen, rango);
            Dictionary<string, double> totales = _hojaResumenController.CrearTablaMetodoLineal(hojaResumen, hojaBase, baremos);
            Dictionary<string, double> propInMasImpMetLineal = _hojaResumenController.CrearTablaTotales(hojaResumen, totales, reclamosValores, baremos, hojaCalXOperario, importeCertificacion);
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



    }
}
