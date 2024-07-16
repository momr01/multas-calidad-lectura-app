using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Interfaces
{
    public interface ICalidadHojaCuadrosController
    {
        void CrearTablaDinEmpleadoTotal(ExcelWorksheet hoja, ExcelRange rango);
        void CrearTablaDinLectorTotal(ExcelWorksheet hoja, ExcelRange rango);
    }
}
