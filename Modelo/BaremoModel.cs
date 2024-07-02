using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Modelo
{
    public class BaremoModel
    {
        private double _t1;
        private double _t2;
        private double _t3;
        private double _alturaT1;
        private double _alturaT3;

        public double T1 { get { return _t1; } set { _t1 = value; } }
        public double T2 { get { return _t2; } set { _t2 = value; } }
        public double T3 { get { return _t3; } set { _t3 = value; } }
        public double AlturaT1 { get { return _alturaT1; } set { _alturaT1 = value; } }
        public double AlturaT3 { get { return _alturaT3; } set { _alturaT3 = value; } }

    }
}
