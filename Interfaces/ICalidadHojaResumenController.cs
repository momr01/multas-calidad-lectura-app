using MultasLectura.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Interfaces
{
    public interface ICalidadHojaResumenController
    {
        void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango);
        Dictionary<string, double> CrearTablaMetodoLineal(ExcelWorksheet hojaResumen, ExcelWorksheet hojaBase, BaremoModel baremos);
        Dictionary<string, double> CrearTablaTotales(ExcelWorksheet hoja, Dictionary<string, double> totales, Dictionary<string, int> reclamos, BaremoModel baremos, ExcelWorksheet hojaCalXOperario, double importeCertificacion);
        void CrearTablaValorFinalMulta(ExcelWorksheet hoja, double propInconformidades, double importeTotalMetLineal, double importeTotalCertificacion, MetaModel metas);
        void CrearTablaBaremosMetas(ExcelWorksheet hoja, BaremoModel baremos, MetaModel metas, double propInconformidades);
    }
}
