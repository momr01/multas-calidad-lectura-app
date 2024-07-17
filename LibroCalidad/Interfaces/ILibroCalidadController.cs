using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Interfaces
{
    public interface ILibroCalidadController
    {
        //void CargarLibroExcel(string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion);
        void GenerarLibroCalidad(string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion, string rutaGuardar);


    }
}
