using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.LibroCalidad.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Controllers
{
    public class CalidadHojaCuadrosController : ICalidadHojaCuadrosController
    {
        private readonly CalidadHojaCuadrosService _service;

        public CalidadHojaCuadrosController()
        {
            _service = new CalidadHojaCuadrosService();
        }

        public void CrearTablaDinEmpleadoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            _service.TablaDinamica(hoja, "A1", rango, "TablaDinEmpleadoTotal", "empleado", "compute_0005", OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum);
        }

        public void CrearTablaDinLectorTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            _service.TablaDinamica(hoja, "d1", rango, "TablaDinLectorTotal", "lector", "nic", OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Count);
        }

    }
}
