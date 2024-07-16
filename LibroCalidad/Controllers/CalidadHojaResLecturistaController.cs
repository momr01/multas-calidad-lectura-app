using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.LibroCalidad.Services;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroCalidad.Controllers
{
    public class CalidadHojaResLecturistaController : ICalidadHojaResLecturistaController
    {
        private readonly CalidadHojaResLecturistaService _service;
        private int _numPrimeraCelda;
        private int _totalInconformidades;
        private int _totalLeidos;
        private double _totalIdeal;

        public CalidadHojaResLecturistaController()
        {
            _service = new CalidadHojaResLecturistaService();
            _numPrimeraCelda = 2;
            _totalInconformidades = 0;
            _totalLeidos = 0;
            _totalIdeal = 0;

        }

        public void CrearTablaLecturistaInconformidades(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            _service.CrearEncabezados(hojaDestino);

            List<EmpleadoModel> empleados = _service.CrearListaEmpleados(hojaCantXOper);

            _service.CalcularInconformidades(hojaCalidadDetalles, ref empleados, ref _totalInconformidades);

            _service.CalcularProporcionIdealLeidos(ref empleados, ref _totalIdeal, ref _totalLeidos);

            List<EmpleadoModel> empleadosOrdenados = empleados.OrderByDescending(emp => emp.Proporcion).ToList();

            List<ColorModel> colores = _service.CargarColores();

            for (int i = 0; i < empleadosOrdenados.Count; i++)
            {
                _service.ColumnaLecturistaA(hojaDestino, _numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaLeidosB(hojaDestino, _numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaInconformidadesC(hojaDestino, _numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaIncXOpD(hojaDestino, _numPrimeraCelda);

                _service.ColumnaIncXNcE(hojaDestino, _numPrimeraCelda, _totalInconformidades);

                _service.ColumnaAcumuladoF(i, hojaDestino, _numPrimeraCelda);

                double ideal = _service.ColumnaIdealG(hojaDestino, _numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaIncXOpIdealH(hojaDestino, _numPrimeraCelda);

                double desvio = _service.ColumnaDesvioI(hojaDestino, _numPrimeraCelda, empleadosOrdenados[i], ideal, _totalIdeal);

                _service.ColorearSegunDesvio(hojaDestino, _numPrimeraCelda, colores, desvio);

                _numPrimeraCelda++;
            }

            _service.CalcularTotal(hojaDestino, 'b', empleados.Count + 2, _totalLeidos);
            _service.CalcularTotal(hojaDestino, 'c', empleados.Count + 2, _totalInconformidades);
            _service.CalcularTotal(hojaDestino, 'g', empleados.Count + 2, (int)Math.Round(_totalIdeal));

            var rangoHojaResLecturista = hojaDestino.Cells[hojaDestino.Dimension.Address];
            LibroExcelHelper.AplicarBordeFinoARango(rangoHojaResLecturista);
        }
    }
}
