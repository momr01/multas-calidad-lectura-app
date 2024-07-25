using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Models
{
    public class DiaMulta
    {
        private int _dia;
        private double _porcentajeIncremento;
        private double _porcentajeObtenido;
        private double _totalMultaDia;

        public int Dia { get { return _dia; } set { _dia = value; } }
        public double PorcentajeIncremento { get { return _porcentajeIncremento; } set { _porcentajeIncremento  = value; } }
        public double PorcentajeObtenido { get { return _porcentajeObtenido; } 
            set
            {
                _porcentajeObtenido = value;
            }  }
        public double TotalMultaDia { get { return _totalMultaDia; }
            set
            {
                _totalMultaDia = value;
            } }

    }

    public class MultaPlazosDia
    {
        private string _tarifa;
        private DiaMulta _1Dia;
        private DiaMulta _2Dia;
        private DiaMulta _3DiaMas;
        private double _porcentajeFueraPlazo;
        private double _importeFueraPlazo;
        private double _totalMulta;
        private string _estadoFinal;
        private ColorModel _colorEstado;

        public string Tarifa { get { return _tarifa; }
            set
            {
                _tarifa = value;
            } }

        public DiaMulta Dia1 { get { return _1Dia; } set {  _1Dia = value; } }
        public DiaMulta Dia2 { get { return _2Dia; }
            set
            {
                _2Dia = value;
            } }
        public DiaMulta Dia3Mas { get { return _3DiaMas; } set {  _3DiaMas = value; } }
        public double PorcentajeFueraPlazo { get { return _porcentajeFueraPlazo; } set { _porcentajeFueraPlazo = value; } }
        public double ImporteFueraPlazo { get { return _importeFueraPlazo; } 
            set
            {
                _importeFueraPlazo = value;
            } }
        public double TotalMulta { get { return _totalMulta; } set { _totalMulta = value; } }
        public string EstadoFinal { get { return _estadoFinal; } set { _estadoFinal = value; } }
        public ColorModel ColorEstado { get { return _colorEstado; } set { _colorEstado = value; } }

        public void DefinirEstadoFinal()
        {
            if(_totalMulta > 0)
            {
                _estadoFinal = "Multa del Período";
                _colorEstado = new ColorModel("multa", Color.Red, Color.White);
            } else
            {
                _estadoFinal = "Saldo a Favor del Período";
                _colorEstado = new ColorModel("saldo a favor", Color.FromArgb(1, 204, 255, 204), Color.Black); ;
            }
        }

        public void CalcularImporteFueraPlazo()
        {
            _importeFueraPlazo = Dia1.TotalMultaDia + Dia2.TotalMultaDia + Dia3Mas.TotalMultaDia;
        }

        public void CalcularTotalAMultar(double bonificacion)
        {
            _totalMulta = _importeFueraPlazo - bonificacion;
        }

       
    }
}
