using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Interfaces
{
    public interface ILibroCalidadController
    {
        void CargarLibroExcel(string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion);
        /*  void CargarLibroCalidadDetalles(string filePath);
          void CargarLibroReclamosDetalles(string filePath);
          void CargarLibroCalidadXOperario(string filePath);
          void CargarBaremos();
          void CargarMetas();*/
        //  void GenerarLibroCalidad(string filePath);

    }
}
