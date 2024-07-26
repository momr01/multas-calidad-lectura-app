using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroPlazos.Controllers
{
    public class LibroPlazosController : ILibroPlazosController
    {
        private readonly IPlazosHojaResumenController _hojaResumenController;
        private readonly BaremoModel _baremos;

        public LibroPlazosController(BaremoModel baremos)
        {
            _baremos = baremos;
            _hojaResumenController = new PlazosHojaResumenController(_baremos);
            
        }

        public void GenerarLibroPlazos(string rutaPlazosDetalles, string rutaGuardar)
        {
            using ExcelPackage libroPlazosDetalles = new(new FileInfo(rutaPlazosDetalles));
            ExcelWorksheet hojaBasePlazosDet = libroPlazosDetalles.Workbook.Worksheets[0];

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroPlazosDetalles.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroPlazosDetalles.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBasePlazosDet.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoPlazosDetalles = hojaBasePlazosDet.Cells[hojaBasePlazosDet.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoPlazosDetalles);

            //Agregar contenido
            AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);
           

            //guardar libro calidad
            libroPlazosDetalles.SaveAs(new FileInfo(rutaGuardar));
        }

        private void AgregarContenidoHojaResumen(ExcelWorksheet hojaResumen,ExcelWorksheet hojaPlazosDetalles, ExcelRange rango)
        {
            int numFila = 1;

            _hojaResumenController.CrearTablaDinCertAtrasoTotal(hojaResumen, rango);
            Dictionary<string, double> importesFinales = _hojaResumenController.CrearTablaDatosPorTarifa(hojaResumen, hojaPlazosDetalles, ref numFila);
            _hojaResumenController.CrearTablaImportesFinales(hojaResumen, importesFinales["multaT1"], importesFinales["multaT2"], ref numFila);

        }
    }
}
