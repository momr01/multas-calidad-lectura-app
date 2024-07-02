using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controlador.Interfaz
{
    public interface ICalidadHojaResumenController
    {
        void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango);
        void CrearTablaMetodoLineal(ExcelWorksheet hojaResumen, ExcelWorksheet hojaBase);
        void CrearTablaTotales(ExcelWorksheet hoja);
        void CrearTablaValorFinalMulta(ExcelWorksheet hoja);
        void CrearTablaBaremosMetas(ExcelWorksheet hoja);
    }
}
