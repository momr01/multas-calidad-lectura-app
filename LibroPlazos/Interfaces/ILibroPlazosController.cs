using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroPlazos.Interfaces
{
    public interface ILibroPlazosController
    {
        void GenerarLibroPlazos(
          string  rutaPlazosDetalles,
          string rutaGuardar
            //string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion, string rutaGuardar
            );
    }
}
