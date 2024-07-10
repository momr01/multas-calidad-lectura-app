using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Models
{
    public class ColorModel
    {
        private string _nombre;
        private Color _fondo;
        private Color _letra;
       
        public string Nombre { get { return _nombre; } set { _nombre = value; } }
        public Color Fondo { get { return _fondo; } set { _fondo = value; } }
        public Color Letra { get { return _letra; } set { _letra = value; } }

        public ColorModel(string Nombre, Color Fondo, Color Letra)
        {
            this.Nombre = Nombre;
            this.Fondo = Fondo;
            this.Letra = Letra;

        }
    }
}
