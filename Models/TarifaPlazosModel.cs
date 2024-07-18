using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Models
{
    public class TarifaPlazosModel
    {
        //'a', "días hábiles t1", "itinerario t1"

        private string _tarifa;
        private string _descripcion;
        private string _encabezado;
        private char _letraInicial;
        private double _baremos;

        public string Tarifa { get { return _tarifa; } set { _tarifa = value; } }
        public string Descripcion { get { return _descripcion;  } set { _descripcion = value; } }
        public string Encabezado { get { return _encabezado; } set { _encabezado = value; } }
        public char LetraInicial { get { return _letraInicial; } 
            set
            {
                _letraInicial = value;
            } }
        public double Baremos { get { return _baremos; } set { _baremos = value; } }

            public TarifaPlazosModel(string Tarifa, string Descripcion, string Encabezado, char LetraInicial, double Baremos)
            {
                this.Tarifa = Tarifa;
                this.Descripcion = Descripcion;
            this.Encabezado = Encabezado;
            this.LetraInicial = LetraInicial;
            this.Baremos = Baremos;

            }
       
    }
}
