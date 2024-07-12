using MultasLectura.Helpers;
using MultasLectura.Interfaces;
using MultasLectura.Models;
using MultasLectura.Services;
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

namespace MultasLectura.Controllers
{
    public class CalidadHojaResLecturistaController : ICalidadHojaResLecturistaController
    {
        private readonly CalidadHojaResLecturistaService _service;
        private int numPrimeraCelda;
        private int totalInconformidades;
        private int totalLeidos;
        private double totalIdeal;

        public CalidadHojaResLecturistaController()
        {
            _service = new CalidadHojaResLecturistaService();
            numPrimeraCelda = 2;
            totalInconformidades = 0;
            totalLeidos = 0;
            totalIdeal = 0;

         }

        public void CrearTablaLecturistaInconformidades(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            _service.CrearEncabezados(hojaDestino);

            List<EmpleadoModel> empleados = _service.CrearListaEmpleados(hojaCantXOper);

            _service.CalcularInconformidades(hojaCalidadDetalles, ref empleados, ref totalInconformidades);

            _service.CalcularProporcionIdealLeidos(ref empleados, ref totalIdeal, ref totalLeidos);

            List<EmpleadoModel> empleadosOrdenados = empleados.OrderByDescending(emp => emp.Proporcion).ToList();

            List<ColorModel> colores = _service.CargarColores();

            for (int i = 0; i < empleadosOrdenados.Count; i++)
            {
                _service.ColumnaLecturistaA(hojaDestino, numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaLeidosB(hojaDestino, numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaInconformidadesC(hojaDestino, numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaIncXOpD(hojaDestino, numPrimeraCelda);

                _service.ColumnaIncXNcE(hojaDestino, numPrimeraCelda, totalInconformidades);

                _service.ColumnaAcumuladoF(i, hojaDestino, numPrimeraCelda);

                double ideal = _service.ColumnaIdealG(hojaDestino, numPrimeraCelda, empleadosOrdenados[i]);

                _service.ColumnaIncXOpIdealH(hojaDestino, numPrimeraCelda);

                double desvio = _service.ColumnaDesvioI(hojaDestino, numPrimeraCelda, empleadosOrdenados[i], ideal, totalIdeal);

                _service.ColorearSegunDesvio(hojaDestino, numPrimeraCelda, colores, desvio);

                numPrimeraCelda++;
            }

            _service.CalcularTotal(hojaDestino, 'b', empleados.Count + 2, totalLeidos);
            _service.CalcularTotal(hojaDestino, 'c', empleados.Count + 2, totalInconformidades);
            _service.CalcularTotal(hojaDestino, 'g', empleados.Count + 2, (int)Math.Round(totalIdeal));

            var rangoHojaResLecturista = hojaDestino.Cells[hojaDestino.Dimension.Address];
            LibroExcelHelper.AplicarBordeFinoARango(rangoHojaResLecturista);
        }
    }
}
