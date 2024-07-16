using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Interfaces
{
    public interface ICalidadHojaResLecturistaController
    {
        void CrearTablaLecturistaInconformidades(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino);
    }
}
