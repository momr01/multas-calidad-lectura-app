using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Modelo
{
    public class Multa
    {
        private int cantidadT1;
        private int cantidadT2;
        private int cantidadT3;
        private int cantidadAlturaT1;
        private int cantidadAlturaT3;
        private double importeT1;
        private double importeT2;
        private double importeT3;
        private double importeAlturaT1;
        private double importeAlturaT3;
        public int CantidadT1 { get { return cantidadT1; } set { cantidadT1 = value; } }
        public int CantidadT2 { get { return cantidadT2; } set { cantidadT2 = value; } }
        public int CantidadT3 { get { return cantidadT3; } set { cantidadT3 = value; } }
        public int CantidadAlturaT1 { get { return cantidadAlturaT1; } set { cantidadAlturaT1 = value; } }
        public int CantidadAlturaT3 { get { return cantidadAlturaT3; } set { cantidadAlturaT3 = value; } }
        public double ImporteT1 { get { return importeT1; } set { importeT1 = value; } }

        public void CalcularImporteT1(double baremo)
        {
            importeT1 = 2 * cantidadT1 * baremo;
        }
    }
}
