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
        void CrearTablaMetodoLineal(ExcelWorksheet hojaResumen, ExcelWorksheet hojaBase, BaremoModel baremos);
        void CrearTablaTotales(ExcelWorksheet hoja);
        void CrearTablaValorFinalMulta(ExcelWorksheet hoja);
        void CrearTablaBaremosMetas(ExcelWorksheet hoja, BaremoModel baremos, MetaModel metas);
    }
}
