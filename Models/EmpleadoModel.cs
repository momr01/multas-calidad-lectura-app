using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Models
{
    public class EmpleadoModel
    {
        private string _nombre;
        private int _leidos;
        private int _inconformidades;
        private double _proporcion;

        public string Nombre { get { return _nombre; } set { _nombre = value; } }
        public int Leidos { get { return _leidos; } set { _leidos = value; } }
        public int Inconformidades { get { return _inconformidades; } set { _inconformidades = value; } }
        public double Proporcion { get { return _proporcion; } set { _proporcion = value; } }

        public void CalcularProporcion()
        {
            _proporcion = (double)_inconformidades / _leidos;
        }

        public EmpleadoModel(string Nombre, int Leidos, int Inconformidades)
        {
            this.Nombre = Nombre;
            this.Leidos = Leidos;
            this.Inconformidades = Inconformidades;

        }
    }
}
