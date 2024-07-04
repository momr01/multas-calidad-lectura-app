using Aspose.Cells.Charts;
using MultasLectura.Helpers;
using MultasLectura.Interfaces;
using MultasLectura.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Controllers
{
    public class MetodoLineal
    {
        private string _descripcion;
        private int _cantidad;
        private double _importe;

        public string Descripcion { get { return _descripcion; } set { _descripcion = value; } }
        public int Cantidad { get { return _cantidad; } set { _cantidad = value; } }
        public double Importe { get { return _importe; } set { _importe = value; } }

        public MetodoLineal(string descripcion, int cantidad, double importe)
        {
            this._descripcion = descripcion;
            this._cantidad = cantidad;
            this._importe = importe;
        }

        //public string Obtener

        public void CalcularImporteConBaremos(double baremo)
        {
           // if (!baremo.Equals(0))
           // {
                _importe = 2 * _cantidad * baremo;
          //  }
           
           // return _importe;
        }

        public void SumarCantidades(int cantidad2)
        {
            _cantidad += cantidad2;

        }

        public void SumarImportes(double importe2)
        {
            _importe += importe2;
        }

    }
    public class CalidadHojaResumenController : ICalidadHojaResumenController
    {

        public void CrearTablaBaremosMetas(ExcelWorksheet hoja, BaremoModel baremos, MetaModel metas, double propInconformidades)
        {
            Dictionary<string, double> datos = new()
            {
                ["T1 y T3"] = baremos.T1,
                ["T2"] = baremos.T2,
                ["Altura T1 y T3"] = baremos.AlturaT1,
                ["Meta"] = metas.Meta1,
                ["Meta 2"] = metas.Meta2,
                ["Obtenido"] = propInconformidades
            };

            var claves = datos.Keys;
            int primeraFilaEstilizar = 2;
            int numFila = 2;

            hoja.Cells["F1"].Value = $"Baremo Lectura desde el {baremos.Fecha}";

            for (int i = 0; i < datos.Count; i++)
            {
                hoja.Cells[$"F{numFila}"].Value = claves.ElementAt(i);
                hoja.Cells[$"G{numFila}"].Value = datos[claves.ElementAt(i)];

                if(numFila >= 5 && numFila <= 7)
                {
                    hoja.Cells[$"G{numFila}"].Style.Numberformat.Format = "0.00%";
                } else
                {
                    LibroExcelHelper.FormatoMoneda(hoja.Cells[$"G{numFila}"]);
                }
                numFila++;
            }
            LibroExcelHelper.AplicarBordeGruesoARango(hoja.Cells[$"F{primeraFilaEstilizar}:G{numFila - 1}"]);
            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"F1:G{numFila - 1}"]);
          
        }

        public void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango)
        {      
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinamicaTipoEstado");
            pivotTable.RowFields.Add(pivotTable.Fields["tipo_certificacion"]);
            pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Count;
        }

        public Dictionary<string, double> CrearTablaMetodoLineal(ExcelWorksheet hojaDestino, ExcelWorksheet hojaOrigen, BaremoModel baremos)
        {
            List<MetodoLineal> datos = new() {
                new MetodoLineal("Certificación Itinerario T1", 0, 0),
                new MetodoLineal("Certificación Itinerario  T2", 0, 0),
                new MetodoLineal("Certificación Itinerario  T3", 0, 0),
                new MetodoLineal("Certificación Itinerario en Altura T1", 0, 0),
                new MetodoLineal("Certificación Itinerario en Altura T3", 0, 0),
            };

            int cantFilas = hojaOrigen.Dimension.Rows;
            int cantCol = hojaOrigen.Dimension.Columns;

            hojaDestino.Cells["A25"].Value = "Método Lineal";

            for (int row = 1; row <= cantFilas; row++)
            {
                for (int col = 1; col <= cantCol; col++)
                {
                    object cellValue = hojaOrigen.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        foreach (MetodoLineal dato in datos)
                        {
                            if(cellValue.ToString() == dato.Descripcion)
                            {
                                dato.Cantidad++;
                            }
                        }
                    }   
                }
            }

            int comienzoTabla = 25;
            int numFila = 26;
            int totalCantidades = 0;
            double totalImportes = 0;

            foreach(MetodoLineal dato in datos)
            {
                if(dato.Descripcion.Contains("T1") || dato.Descripcion.Contains("T3"))
                {
                    if (dato.Descripcion.Contains("Altura"))
                    {
                        dato.CalcularImporteConBaremos(baremos.AlturaT1);
                    } else
                    {
                        dato.CalcularImporteConBaremos(baremos.T1);
                    }
                } else
                {
                    dato.CalcularImporteConBaremos(baremos.T2);
                }
                
                totalCantidades += dato.Cantidad;
                totalImportes += dato.Importe;

                hojaDestino.Cells[$"A{numFila}"].Value = dato.Descripcion;
                hojaDestino.Cells[$"B{numFila}"].Value = dato.Cantidad;
                hojaDestino.Cells[$"C{numFila}"].Value = dato.Importe;

                numFila++;
            }

            hojaDestino.Cells[$"B{numFila}"].Value = totalCantidades;
            hojaDestino.Cells[$"C{numFila}"].Value = totalImportes;
            
            LibroExcelHelper.FondoSolido(hojaDestino.Cells[$"C{numFila}"], Color.FromArgb(1, 252, 213, 180));
            LibroExcelHelper.FormatoMoneda(hojaDestino.Cells[$"C{comienzoTabla + 1}:C{numFila}"]);
            LibroExcelHelper.AplicarBordeFinoARango(hojaDestino.Cells[$"A{comienzoTabla}:C{numFila}"]);

            return new() {
                ["total"] = totalCantidades,
                ["importe"] = totalImportes
            };

        }

        public Dictionary<string, double> CrearTablaTotales(ExcelWorksheet hoja, Dictionary<string, double> totales, Dictionary<string, int> reclamos, BaremoModel baremos, ExcelWorksheet hojaCalXOperario, double importeCertificacion)
        {
            int totalCertificado = LibroExcelHelper.SumarColumnaInt(hojaCalXOperario, 5, 2);

            List<MetodoLineal> datos = new() {
                new MetodoLineal("Anomalias de Facturacion NC", int.Parse(totales["total"].ToString()), totales["importe"]),
                new MetodoLineal("Reclamos procedentes T1", reclamos["t1"], 0),
                new MetodoLineal("Reclamos procedentes T2", reclamos["t2"], 0),
                new MetodoLineal("Total de NC por Metodo Lineal (0,15% al 0,3%)", int.Parse(totales["total"].ToString()), totales["importe"]),
                new MetodoLineal("Totales Certificado", totalCertificado, importeCertificacion),
            };

            int totalCantidadesReclamos = reclamos["t1"] + reclamos["t2"];
            double totalImportesReclamos = 0;
            int numFila = 35;
            int filaInicial = 35;

            hoja.Cells[$"A{numFila}"].Value = "Descripción";
            hoja.Cells[$"B{numFila}"].Value = "TOTAL";
            hoja.Cells[$"C{numFila}"].Value = "IMPORTE";

            foreach (MetodoLineal dato in datos)
            {
                if (dato.Descripcion.ToLower().Contains("reclamos"))
                {
                    if (dato.Descripcion.ToLower().Contains("t1"))
                    {
                        dato.CalcularImporteConBaremos(baremos.T1);
                    } else
                    {
                        dato.CalcularImporteConBaremos(baremos.T2);
                    }

                    totalImportesReclamos += dato.Importe;
                }

            }

      

            foreach (MetodoLineal dato in datos)
            {

                if(dato.Descripcion.ToLower().Contains("metodo lineal"))
                {
                    dato.SumarCantidades(totalCantidadesReclamos);
                    dato.SumarImportes(totalImportesReclamos);  
                }

                hoja.Cells[$"A{numFila + 1}"].Value = dato.Descripcion;
                hoja.Cells[$"B{numFila + 1}"].Value = dato.Cantidad;
                hoja.Cells[$"C{numFila + 1}"].Value = dato.Importe;

                numFila++;


            }

            double propInconformidades = (double)datos.Where(dato => dato.Descripcion.ToLower().Contains("metodo lineal")).FirstOrDefault().Cantidad / datos.Where(dato => dato.Descripcion.ToLower().Contains("certificado")).FirstOrDefault().Cantidad;

            hoja.Cells[$"D{filaInicial + datos.Count}"].Value = propInconformidades;

            //hoja.Cells[$"G{numFila}"].Style.Numberformat.Format = "0.00%";
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"D{filaInicial + datos.Count}"]);

            LibroExcelHelper.AplicarBordeFinoARango(hoja.Cells[$"A{filaInicial}:C{filaInicial + datos.Count}"]);
            LibroExcelHelper.FormatoMoneda(hoja.Cells[$"C{filaInicial + 1}:C{filaInicial + datos.Count}"]);
            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"B{filaInicial + datos.Count}:C{filaInicial + datos.Count}"]);

            hoja.Cells.AutoFitColumns();

            return new() { 
                
               ["propInconformidades"] = propInconformidades,
                ["totalMetLineal"] = datos.Where(dato => dato.Descripcion.ToLower().Contains("metodo lineal")).FirstOrDefault().Importe
            };
        }

        public void CrearTablaValorFinalMulta(ExcelWorksheet hoja, double propInconformidades, double importeTotalMetLineal, double importeTotalCertificacion, MetaModel metas)
        {
            double importeMultaFinal = 0;

            if (propInconformidades > metas.Meta1)
            {
                if(propInconformidades > metas.Meta2)
                {
                    double calcAuxiliar = (propInconformidades - metas.Meta1) / (0.01 - metas.Meta1);
                    importeMultaFinal = importeTotalCertificacion * Math.Pow(calcAuxiliar, 2);
                   
                } else
                {
                    importeMultaFinal = (propInconformidades * importeTotalMetLineal) / propInconformidades;
                }
            }

            double propMultaSobreTotalCert = importeMultaFinal / importeTotalCertificacion;






            // hoja.Cells["A44"].RichText.Add("Multa").Bold = true;
            hoja.Cells["A44"].Value = "Multa";
            hoja.Cells["B44"].Value = importeMultaFinal;
            hoja.Cells["C44"].Value = propMultaSobreTotalCert;

            //  hoja.Cells["A44:B44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            // hoja.Cells["A44:B44"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

            LibroExcelHelper.FormatoNegrita(hoja.Cells["A44:B44"]);
            LibroExcelHelper.FondoSolido(hoja.Cells["A44:B44"], Color.FromArgb(1, 255, 192, 0));
            LibroExcelHelper.FormatoMoneda(hoja.Cells["B44"]);
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells["C44"]);
            LibroExcelHelper.AplicarBordeFinoARango(hoja.Cells["A44:B44"]);

            hoja.Cells.AutoFitColumns();

        }
    }
}
