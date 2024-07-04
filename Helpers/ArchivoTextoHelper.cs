using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MultasLectura.Models;

namespace MultasLectura.Helpers
{
    public class ArchivoTextoHelper
    {
        public static void VerificarExisteArchivoBaremos(BaremoModel _baremos)
        {
            string pathProyecto = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(pathProyecto, "baremos.txt");

            if (File.Exists(filePath))
            {
                LeerArchivoBaremos(filePath, _baremos);
            }
            else
            {
                CrearArchivoBaremos(filePath);
                LeerArchivoBaremos(filePath, _baremos);
            }
        }

        public static void VerificarExisteArchivoMetas(MetaModel _metas)
        {
            string pathProyecto = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(pathProyecto, "metas.txt");

            if (File.Exists(filePath))
            {
                LeerArchivoMetas(filePath, _metas);
            }
            else
            {
                CrearArchivoMetas(filePath);
                LeerArchivoMetas(filePath, _metas);
            }
        }

        private static void LeerArchivoBaremos(string filePath, BaremoModel _baremos)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                string linea;

                while ((linea = reader.ReadLine()) != null)
                {
                    var arregloLinea = linea.Split(';');
                    switch (arregloLinea[0])
                    {
                        case "t1":
                            _baremos.T1 = double.Parse(arregloLinea[1]);
                            break;
                        case "t2":
                            _baremos.T2 = double.Parse(arregloLinea[1]);
                            break;
                        case "t3":
                            _baremos.T3 = double.Parse(arregloLinea[1]);
                            break;
                        case "alturat1":
                            _baremos.AlturaT1 = double.Parse(arregloLinea[1]);
                            break;
                        case "alturat3":
                            _baremos.AlturaT3 = double.Parse(arregloLinea[1]);
                            break;
                        case "fecha":
                            _baremos.Fecha = arregloLinea[1];
                            break;

                    }
                }
            }

        }

        private static void LeerArchivoMetas(string filePath, MetaModel _metas)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                string linea;

                while ((linea = reader.ReadLine()) != null)
                {
                    var arregloLinea = linea.Split(';');
                    switch (arregloLinea[0])
                    {
                        case "meta1":
                            _metas.Meta1 = double.Parse(arregloLinea[1]);
                            break;
                        case "meta2":
                            _metas.Meta2 = double.Parse(arregloLinea[1]);
                            break;

                    }
                }
            }

        }

        private static void CrearArchivoBaremos(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.WriteLine("fecha;01/02/2024");
                writer.WriteLine("t1;0");
                writer.WriteLine("t2;0");
                writer.WriteLine("t3;0");
                writer.WriteLine("alturat1;0");
                writer.WriteLine("alturat3;0");
            }
        }

        private static void CrearArchivoMetas(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.WriteLine("meta1;0");
                writer.WriteLine("meta2;0");
            }
        }


    }
}
