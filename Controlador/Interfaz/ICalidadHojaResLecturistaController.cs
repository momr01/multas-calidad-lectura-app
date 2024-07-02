using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controlador.Interfaz
{
    public interface ICalidadHojaResLecturistaController
    {
        void CrearTablaLecturistaInconformidades(ExcelWorksheet hoja, ExcelRange rango);
    }
}
