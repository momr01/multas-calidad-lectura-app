using Aspose.Cells.Drawing;
using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
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

            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}1", tarifa.Encabezado.ToUpper(), Enums.TipoOpCelda.Value);
          //  hojaResumen.Cells[$"{primLetra}1:{cuarLetra}1"].Merge = true;
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}2", "FTL", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{segLetra}2", "k", Enums.TipoOpCelda.Value);

            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}1:{cuarLetra}1");

            // hojaResumen.Cells[$"{primLetra}1"].Value = tarifa.Encabezado.ToUpper();

           
            //hojaResumen.Cells[$"{primLetra}2"].Value = "FTL";
            
           // hojaResumen.Cells[$"{segLetra}2"].Value = "k";
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}2:{segLetra}2"], Color.FromArgb(1, 204, 255, 255 ));
           // hojaResumen.Cells[$"{tercLetra}2"].Value = "Qij";
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}2", "Qij", Enums.TipoOpCelda.Value);
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}2:{cuarLetra}2");
           // hojaResumen.Cells[$"{tercLetra}2:{cuarLetra}2"].Merge = true;
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{tercLetra}2"], Color.FromArgb(1, 255, 0, 0));

            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}1:{cuarLetra}2"]);

            List<int> ftlLista = new()
          ;

           

            for (int i = -14; i < 18; i++)
            {
                ftlLista.Add(i);

            }


            int totalCantidades = 0;
            int totalEnPlazo = 0;
            int totalFueraDePlazo = 0;
            int totalK1 = 0;
            double importeK1 = 0;
            int totalK2 = 0;
            double importeK2 = 0;
            int totalK3oMas = 0;
            double importeK3oMas = 0;



            foreach (int ftl in ftlLista)
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", ftl, Enums.TipoOpCelda.Value);
               // hojaResumen.Cells[$"{primLetra}{numFila}"].Value = ftl;

                int cantidad = CalcularCantAtrasos(hojaReclDetalles, ftl, tarifa.Descripcion);
                totalCantidades += cantidad;

               // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = cantidad;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", cantidad, Enums.TipoOpCelda.Value);


                ColorearRangoSegunNum(ftl, hojaResumen, numFila, primLetra, tercLetra);

                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, tarifa.LetraInicial);
                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra2);
                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra3);

                CargarDatosColK(ftl, hojaResumen, numFila, segLetra);

               // hojaResumen.Cells[$"{tercLetra}{numFila}:{cuarLetra}{numFila}"].Merge = true;
                LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");

                if (ftl >= -3 && ftl <= 1)
                {
                    totalEnPlazo += cantidad;
                } else
                {
                    totalFueraDePlazo += cantidad;
                }

                if(ftl == -4 || ftl == 2)
                {
                    totalK1 += cantidad;
                    //importeK1 += (double)cantidad * int.Parse(hojaResumen.Cells[$"{segLetra}{numFila}"].Value.ToString());
                    importeK1 += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                } else if(ftl == -5 || ftl == 3)
                {
                    totalK2 += cantidad;
                    importeK2 += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                } else if(ftl <= -6 || ftl >= 4)
                {
                    totalK3oMas += cantidad;
                    importeK3oMas += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                }




                numFila++;

            }

            // LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"A{numFilaInicial}:A{numFila}"]);
         //   hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
          //  hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            LibroExcelHelper.CentrarContenidoCelda(hojaResumen, $"{primLetra}1:{cuarLetra}{numFila - 1}");

            /* List<string> headersTabla2 = new()
             {
                 "Certificado", "Valor de Lectura", "Dentro Plazo", "Bonificación"
             };*/

            double porcentajeDentroPlazo = (double)totalEnPlazo / totalCantidades;

            double bonifica = porcentajeDentroPlazo >= 0.7 ? 1 : 0;
            double totalImpCert = (double)totalCantidades * tarifa.Baremos;

            List<double> dentroPlazo = new() { porcentajeDentroPlazo, double.Parse(totalEnPlazo.ToString()) };
            List<double> bonificacion = new() { bonifica,  0 };

            Dictionary<string, Generics> valoresTabla2 = new()
            {
                { "Certificado", new Generics( totalImpCert) },
                { "Valor de Lectura", new Generics(tarifa.Baremos) },
                { "Dentro Plazo", new Generics(dentroPlazo) },
                { "Bonificación", new Generics(bonificacion) }
            };

            double importeBonif = 0;

            var keys = valoresTabla2.Keys;

            for (int  i = 0; i < valoresTabla2.Count; i++)
            {
                var valor = valoresTabla2[keys.ElementAt(i)].Value;

                // hojaResumen.Cells[$"{primLetra}{numFila}"].Value = headersTabla2[i];
              //  hojaResumen.Cells[$"{primLetra}{numFila}"].Value = keys.ElementAt(i);
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", keys.ElementAt(i), Enums.TipoOpCelda.Value);

                if (valor is double)
                {
                    //double valorD = double.Parse(valor.ToString()!);
                    //  hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = valoresTabla2[keys.ElementAt(i)].GetValue<double>();
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", valoresTabla2[keys.ElementAt(i)].GetValue<double>(), 
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");
                } else if(valor is int)
                {
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", valoresTabla2[keys.ElementAt(i)].GetValue<int>(), 
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");
                } else if(valor is List<double>)
                {
                    List<double> lista = valoresTabla2[keys.ElementAt(i)].GetValue<List<double>>();
                    if (i == 3)
                    {
                        string textoBonifica = "No Bonifica";
                      //  double importeBonif = 0;
                        ColorModel colores = new("no bonifica", Color.Red, Color.White);

                        if (lista[0] == 1)
                        {
                            textoBonifica = "Bonifica";
                            importeBonif = (0.1 * totalImpCert * ((porcentajeDentroPlazo - 0.7)/0.3));
                            colores.Nombre = textoBonifica;
                            colores.Fondo = Color.FromArgb(1, 204, 255, 204);
                            colores.Letra = Color.Black;
                        } 

                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", textoBonifica,
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", importeBonif,
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.ColorFondoLetra(hojaResumen, letra3, numFila, colores);

                    } else
                    {
                        
                        // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = lista[0];
                        //hojaResumen.Cells[$"{cuarLetra}{numFila}"].Value = lista[1];
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", lista[0],
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", lista[1],
                           Enums.TipoOpCelda.Value);
                    }
                   

                } else
                {
                    // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = "holiiis";
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", 0,
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");

                }
               
              //  hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"].Merge = true;
                LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}{numFila}:{segLetra}{numFila}");

                if(i == 0 || i==1)
                {
                    LibroExcelHelper.FormatoMoneda(hojaResumen.Cells[$"{tercLetra}{numFila}"]);

                } else if(i== 2)
                {
                    LibroExcelHelper.FormatoPorcentaje(hojaResumen.Cells[$"{tercLetra}{numFila}"]);
                    //LibroExcelHelper.MIles
                } else
                {
                    LibroExcelHelper.FormatoMoneda(hojaResumen.Cells[$"{cuarLetra}{numFila}"]);
                }
                numFila++;
            }



            double porcentajeUnDia = 0.02;
            double porcentajeDosDias = porcentajeUnDia * 2;
            double porcentajeMasTresDias = porcentajeDosDias * 2.5;




            MultaPlazosDia tabla3 = new()
            {
                Tarifa = "t1",
               // ImporteFueraPlazo = 0,
                PorcentajeFueraPlazo = porcentajeUnDia + porcentajeDosDias + porcentajeMasTresDias,
              //  TotalMulta = 7889,

                Dia1 = new()
                {
                    Dia = 1,
                    PorcentajeIncremento = porcentajeUnDia,
                    PorcentajeObtenido = (double)totalK1 / totalCantidades,
                    // TotalMultaDia = (porcentajeUnDia * tarifa.Baremos) * importeK1,
                    TotalMultaDia = CalcularTotalMultaDia(porcentajeUnDia, tarifa.Baremos, importeK1),
                },

                 Dia2 = new()
                 {
                     Dia = 2,
                     PorcentajeIncremento = porcentajeDosDias,
                     PorcentajeObtenido = (double)totalK2/totalCantidades ,
                     TotalMultaDia = CalcularTotalMultaDia(porcentajeDosDias, tarifa.Baremos, importeK2),
                 },

                  Dia3Mas = new()
                  {
                      Dia = 3,
                      PorcentajeIncremento = porcentajeMasTresDias,
                      PorcentajeObtenido = (double)totalK3oMas/totalCantidades,
                      TotalMultaDia = CalcularTotalMultaDia(porcentajeMasTresDias, tarifa.Baremos, importeK3oMas),
                  }
            };

            // Obtener el tipo de la instancia
            Type tipo = tabla3.GetType();

            // Obtener todas las propiedades de la instancia
            PropertyInfo[] propiedades = tipo.GetProperties();

            // Recorrer las propiedades y sus valores

            numFila++;
            List<char> letras4 = new() { tarifa.LetraInicial, letra2, letra3, letra4 };

            foreach (PropertyInfo propiedad in propiedades)
            {
               // if(propiedad.PropertyType == typeof(int))


                string nombre = propiedad.Name;
                object valor = propiedad.GetValue(tabla3)!;


                // MessageBox.Show(nombre);
                // object value = property.GetValue(person);
                // Console.WriteLine($"{name}: {value}");

                if (valor != null && valor.GetType() == typeof(DiaMulta))
                {
                    Type diaTipo = valor.GetType();
                    PropertyInfo[] diaProps = diaTipo.GetProperties();

                    foreach (PropertyInfo prop in diaProps)
                    {
                        //string nombreDia = prop.Name;
                       // object addressValue = prop.GetValue(valor);
                       // Console.WriteLine($"  {addressName}: {addressValue}");
                    }
                    for (int i = 0; i < diaProps.Length; i++)
                    {
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letras4[i]}{numFila}", diaProps[i].GetValue(valor), Enums.TipoOpCelda.Value);
                       
                    }

                    numFila++;
                } 



               /* if (nombre.ToLower().Contains("dia"))
                {
                    // Obtener el tipo de la instancia
                    Type tipoDia = propiedad.GetType();

                    // Obtener todas las propiedades de la instancia
                    PropertyInfo[] propsDia = tipoDia.GetProperties();

                  //  foreach(PropertyInfo prop in propsDia)
                  //  {
                   //     LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{}{fila}", valor, Enums.TipoOpCelda.Value);
                   // }
                   for(int i = 0; i < propsDia.Length; i++)
                    {
                        string jjj = propsDia[i].GetValue(tipoDia)!.ToString()!;
                        MessageBox.Show(jjj);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letras4[i]}{numFila}", 0, Enums.TipoOpCelda.Value);
                    }


                   // FilaDiaTabla3(hojaResumen, new() { tarifa.LetraInicial, letra2, letra3, letra4}, numFila, 5);
                }*/

                
            }

            tabla3.CalcularImporteFueraPlazo();
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", "Total Fuera Plazo", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", tabla3.PorcentajeFueraPlazo, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", tabla3.ImporteFueraPlazo, Enums.TipoOpCelda.Value);

            tabla3.CalcularTotalAMultar(importeBonif);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 1}", "Total a Multar", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila + 1}", tabla3.TotalMulta, Enums.TipoOpCelda.Value);

            tabla3.DefinirEstadoFinal();
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila + 2}", tabla3.EstadoFinal, Enums.TipoOpCelda.Value);
            LibroExcelHelper.ColorFondoLetra(hojaResumen, letra3, numFila + 2, tabla3.ColorEstado);


            //  for( int i = 0; i < 5; i++)
            //{


            //}
           // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 6}", "FUERA DE PLAZO DEL PERIODO", Enums.TipoOpCelda.Value);
           // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 10}", "BONIFICACION DEL PERIODO", Enums.TipoOpCelda.Value);







        }

        private double CalcularImporteK(int cantidad, ExcelWorksheet hoja, string letra, int fila)
        {
            return (double)cantidad * int.Parse(hoja.Cells[$"{letra}{fila}"].Value.ToString()!);
        }

        private double CalcularTotalMultaDia(double porcentaje, double baremos, double importeK)
        {
            return (porcentaje * baremos) * importeK;
        }

        private void FilaDiaTabla3(ExcelWorksheet hoja, List<char> letras, int fila, double valor)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[0]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[1]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[2]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[3]}{fila}", valor, Enums.TipoOpCelda.Value);
        }

        private void CargarDatosColK(int ftl, ExcelWorksheet hojaResumen, int numFila, string letra)
        {
            if (ftl <= -3)
            {
               // hojaResumen.Cells[$"{letra}{numFila}"].Value = (ftl + 3) * -1;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", (ftl + 3) * -1,
                         Enums.TipoOpCelda.Value);

            }
            else if (ftl >= 2)
            {
               // hojaResumen.Cells[$"{letra}{numFila}"].Value = ftl - 1;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", ftl - 1,
                        Enums.TipoOpCelda.Value);
            }
            else
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", 0, Enums.TipoOpCelda.Value);
            }
        }

        private void AplicarBordeParcialARango(int numFila, int numFilaInicial, List<int> ftl, ExcelWorksheet hojaResumen, char letraCelda)
        {
            string letra = letraCelda.ToString().ToUpper();
            string rango = $"{letra}{numFila}";

            if (numFila == numFilaInicial)
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.SupDerIzq);
                /*hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

            }
            else if (numFila == (ftl.Count + numFilaInicial - 1))
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.InfDerIzq);
                /*hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

            }
            else
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.CentroDerIzq);

              /*  hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/
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
