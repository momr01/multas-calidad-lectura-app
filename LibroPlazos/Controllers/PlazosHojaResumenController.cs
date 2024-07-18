using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace MultasLectura.LibroPlazos.Controllers
{
    public class Generics
    {
        public object Value { get; set; }

        public Generics(object value)
        {
            Value = value;
        }

        public T GetValue<T>()
        {
            return (T)Value;
        }
    }


    public class PlazosHojaResumenController : IPlazosHojaResumenController
    {
        private readonly BaremoModel _baremos;

        public PlazosHojaResumenController(BaremoModel baremos)
        {
            _baremos = baremos;
        }

        public void CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles)
        {
            List<TarifaPlazosModel> tarifas = new()
            {
                new("t1", "itinerario t1", "días hábiles t1", 'a', _baremos.T1),
                new("t2", "itinerario  t2", "días hábiles t2", 'f', _baremos.T2),
                new("t3", "itinerario  t3", "días hábiles t3", 'k', _baremos.T3),
                new("altura t1", "altura t1", "días hábiles altura t1", 'p', _baremos.AlturaT1),
                new("altura t3", "altura t3", "días hábiles altura t3", 'u', _baremos.AlturaT3)
            };

            foreach(TarifaPlazosModel tarifa in tarifas)
            {
                CrearTablaDatosPorTarifa(hojaResumen, hojaReclDetalles, tarifa);
            }
           



        }

        private char ObtenerLetraSiguiente(char letra)
        {
            // Verificar si la letra es una letra del alfabeto
           // if (!char.IsLetter(letra))
           // {
          //      throw new ArgumentException("El carácter proporcionado no es una letra.");
          //  }

            // Obtener el código ASCII de la letra
            int asciiCode = (int)letra;

            // Incrementar el código ASCII y ajustar para el caso de 'z' y 'Z'
            if (letra == 'z')
            {
                asciiCode = (int)'a';
            }
            else if (letra == 'Z')
            {
                asciiCode = (int)'A';
            }
            else
            {
                asciiCode++;
            }

            // Convertir de nuevo a carácter
            return (char)asciiCode;
        }

        private void CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, TarifaPlazosModel tarifa )
        {
            int numFilaInicial = 3;
            int numFila = 3;

            string primLetra = tarifa.LetraInicial.ToString().ToUpper();
            char letra2 = ObtenerLetraSiguiente(tarifa.LetraInicial);
            string segLetra = letra2.ToString().ToUpper();
            char letra3 = ObtenerLetraSiguiente(letra2);
            string tercLetra = letra3.ToString().ToUpper();
            char letra4 = ObtenerLetraSiguiente(letra3);
            string cuarLetra = letra4.ToString().ToUpper();

            hojaResumen.Cells[$"{primLetra}1"].Value = tarifa.Encabezado.ToUpper();

            hojaResumen.Cells[$"{primLetra}1:{cuarLetra}1"].Merge = true;
            hojaResumen.Cells[$"{primLetra}2"].Value = "FTL";
            
            hojaResumen.Cells[$"{segLetra}2"].Value = "k";
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}2:{segLetra}2"], Color.FromArgb(1, 204, 255, 255 ));
            hojaResumen.Cells[$"{tercLetra}2"].Value = "Qij";
            hojaResumen.Cells[$"{tercLetra}2:{cuarLetra}2"].Merge = true;
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{tercLetra}2"], Color.FromArgb(1, 255, 0, 0));

            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}1:{cuarLetra}2"]);

            List<int> ftlLista = new()
          ;

           

            for (int i = -14; i < 18; i++)
            {
                ftlLista.Add(i);

            }

          



            foreach (int ftl in ftlLista)
            {
                hojaResumen.Cells[$"{primLetra}{numFila}"].Value = ftl;

                int cantidad = CalcularCantAtrasos(hojaReclDetalles, ftl, tarifa.Descripcion);

                hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = cantidad;

                ColorearRangoSegunNum(ftl, hojaResumen, numFila, primLetra, tercLetra);

                AplicarBordeARango(numFila, numFilaInicial, ftlLista, hojaResumen, tarifa.LetraInicial);
                AplicarBordeARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra2);
                AplicarBordeARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra3);

                CargarDatosColK(ftl, hojaResumen, numFila, segLetra);

                hojaResumen.Cells[$"{tercLetra}{numFila}:{cuarLetra}{numFila}"].Merge = true;




                numFila++;

            }

            // LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"A{numFilaInicial}:A{numFila}"]);
            hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            /* List<string> headersTabla2 = new()
             {
                 "Certificado", "Valor de Lectura", "Dentro Plazo", "Bonificación"
             };*/

            List<double> dentroPlazo = new() { 41, 105000 };

            Dictionary<string, Generics> valoresTabla2 = new();
            valoresTabla2.Add("Certificado", new Generics(78000000));
            valoresTabla2.Add("Valor de Lectura", new Generics(tarifa.Baremos));
            valoresTabla2.Add("Dentro Plazo", new Generics(dentroPlazo));
            valoresTabla2.Add("Bonificación", new Generics(566));


            for (int  i = 0; i < 4; i++)
            {
               // hojaResumen.Cells[$"{primLetra}{numFila}"].Value = headersTabla2[i];
                hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"].Merge = true;
                numFila++;
            }

          

        }

        private void CargarDatosColK(int ftl, ExcelWorksheet hojaResumen, int numFila, string letra)
        {
            if (ftl <= -3)
            {
                hojaResumen.Cells[$"{letra}{numFila}"].Value = (ftl + 3) * -1;
            }
            else if (ftl >= 2)
            {
                hojaResumen.Cells[$"{letra}{numFila}"].Value = ftl - 1;
            }
            else
            {
                hojaResumen.Cells[$"{letra}{numFila}"].Value = 0;
            }
        }

        private void AplicarBordeARango(int numFila, int numFilaInicial, List<int> ftl, ExcelWorksheet hojaResumen, char letraCelda)
        {
            string letra = letraCelda.ToString().ToUpper();

            if (numFila == numFilaInicial)
            {
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

            }
            else if (numFila == (ftl.Count + numFilaInicial - 1))
            {
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

            }
            else
            {

                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            }
        }

        private void ColorearRangoSegunNum(int num, ExcelWorksheet hojaResumen, int numFilaInicial, string letra1, string letra3)
        {
            if (num <= -6 || num >= 4)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 102, 0));
            }
            else if (num == -5 || num == 3)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 204, 153));
            }
            else if (num == -4 || num == 2)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 255, 153));
            }
            else
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 204, 255, 204));
            }
            /* List<int> ftl = new()
          ;

             for (int i = -14; i < 18; i++)
             {
                 ftl.Add(i);

             }

             foreach (int num in ftl)
             {
                if(num <= -6 || num >= 4)
                 {
                     LibroExcelHelper.FondoSolido();
                 }

             }*/



        }



        private int CalcularCantAtrasos(ExcelWorksheet hojaBase, int atraso, string tipo)
        {
            int cantFilas = hojaBase.Dimension.Rows;
          

            int colTipItin = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "tip_itin");
            int colAtraso = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "atraso");
            int colCantSum = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "cant_sum");

            int cantFinal = 0;

            if (colTipItin != -1 && colAtraso != -1 && colCantSum != -1)
            {
                for (int fila = 1; fila <= cantFilas; fila++)
                {
                    object cellValue = hojaBase.Cells[fila, colTipItin].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.ToLower().Contains(tipo))
                        {
                            int atrasoBase = int.Parse(hojaBase.Cells[fila, colAtraso].Value.ToString()!);

                            if(atrasoBase == atraso)
                            {
                                cantFinal += int.Parse(hojaBase.Cells[fila, colCantSum].Value.ToString()!);
                            }
                           


                        }
                    }
                }
            }

            return cantFinal;
        }

       

        public void CrearTablaDinCertAtrasoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["Z1"], rango, "TablaDinamicaCertAtrasoTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["tip_itin"]);
            var atraso = pivotTable.RowFields.Add(pivotTable.Fields["atraso"]);
            pivotTable.DataFields.Add(pivotTable.Fields["cant_sum"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Sum;

            atraso.Sort = eSortType.Ascending;
        }

        public void CrearTablaImportesFinales()
        {
            throw new NotImplementedException();
        }


    }
}
