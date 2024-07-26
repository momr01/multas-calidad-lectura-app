using MultasLectura.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroPlazos.Interfaces
{
    public interface IPlazosHojaResumenController
    {
        void CrearTablaDinCertAtrasoTotal(ExcelWorksheet hoja, ExcelRange rango);
        Dictionary<string, double> CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, ref int numFila);
        void CrearTablaImportesFinales(ExcelWorksheet hoja, double importeT1, double importeT2, ref int numFila);
    }
}
