using Aspose.Cells;
using Aspose.Cells.Charts;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Runtime.CompilerServices;
using NPOI.SS.Formula.Functions;
using MultasLectura.Interfaces;
using MultasLectura.Models;
using MultasLectura.Controllers;
using MultasLectura.Views;
using MultasLectura.Helpers;

namespace MultasLectura
{
    public partial class GenerarLibroCalidad : Form
    {
        private readonly ILibroCalidadController _calidadController;
        // private readonly Action<string> _cargarLibroExcelFuncion;
        private readonly BaremoModel _baremos = new BaremoModel();
        private readonly MetaModel _metas = new MetaModel();

        private Loader loaderForm;

        public GenerarLibroCalidad()
        {
            InitializeComponent();
            ArchivoTextoHelper.VerificarExisteArchivoBaremos(_baremos);
            ArchivoTextoHelper.VerificarExisteArchivoMetas(_metas);
            _calidadController = new LibroCalidadController(_baremos!, _metas!);
            // _cargarLibroExcelFuncion = _calidadController.CargarLibroExcel;
            loaderForm = new Loader();

        }


        private void btnCalidadDetalles_Click(object sender, EventArgs e)
        {
            // LibroExcelModel.IniciarProcesoCarga(txtRutaCalidadDetalles, _cargarLibroExcelFuncion);
            LibroExcelHelper.IniciarProcesoCarga(txtRutaCalidadDetalles);
        }

        private void btnReclamosDetalles_Click(object sender, EventArgs e)
        {
            //  LibroExcelModel.IniciarProcesoCarga(txtRutaReclamosDetalles, _cargarLibroExcelFuncion);
            LibroExcelHelper.IniciarProcesoCarga(txtRutaReclamosDetalles);
            /* string filePath = LibroExcelModel.CargarLibroExcel();

             if (string.IsNullOrEmpty(filePath))
             {
                 txtRutaReclamosDetalles.Text = string.Empty;
                 LibroExcelModel.MostrarMensaje("Ocurrió un error al intentar cargar el archivo. Por favor inténtelo nuevamente", true);
             }
             else
             {
                 txtRutaReclamosDetalles.Text = filePath;
             }*/
        }

        private void btnCalXOperarios_Click(object sender, EventArgs e)
        {
            // LibroExcelModel.IniciarProcesoCarga(txtRutaCalXOperarios, _cargarLibroExcelFuncion);
            LibroExcelHelper.IniciarProcesoCarga(txtRutaCalXOperarios);
            /* string filePath = LibroExcelModel.CargarLibroExcel();

             if (string.IsNullOrEmpty(filePath))
             {
                 txtRutaCalXOperarios.Text = string.Empty;
                 LibroExcelModel.MostrarMensaje("Ocurrió un error al intentar cargar el archivo. Por favor inténtelo nuevamente", true);
             }
             else
             {
                 txtRutaCalXOperarios.Text = filePath;
             }*/
        }

        /*
        private void CrearArchivoBaremos(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.WriteLine("t1;0");
                writer.WriteLine("t2;0");
                writer.WriteLine("t3;0");
                writer.WriteLine("alturat1;0");
                writer.WriteLine("alturat3;0");
            }
        }

        private void CrearArchivoMetas(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.WriteLine("meta1;0");
                writer.WriteLine("meta2;0");
            }
        }
        */

        /*
        private void LeerArchivoBaremos(string filePath)
        {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string linea;

                    while ((linea = reader.ReadLine()) != null)
                    {
                        var arregloLinea = linea.Split(';');
                    switch(arregloLinea[0])
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

                    }
                    }
                }
           
        }
        */

        /* private void LeerArchivoMetas(string filePath)
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

         }*/

        /*

        private void VerificarExisteArchivoBaremos()
        {
            string pathProyecto = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(pathProyecto, "baremos.txt");

            if (File.Exists(filePath))
            {
                // LeerArchivoBaremos(filePath);
                ArchivoTextoModel.LeerArchivoBaremos(filePath, _baremos);
            }
            else
            {
                CrearArchivoBaremos(filePath);
                // LeerArchivoBaremos(filePath);
                ArchivoTextoModel.LeerArchivoBaremos(filePath, _baremos);
            }
        }

        private void VerificarExisteArchivoMetas()
        {
            string pathProyecto = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(pathProyecto, "metas.txt");

            if (File.Exists(filePath))
            {
                LeerArchivoMetas(filePath);
            }
            else
            {
                CrearArchivoMetas(filePath);
                LeerArchivoMetas(filePath);
            }
        }
        */

        private void CargarDatosBaremos()
        {
            baremosT1.Text = _baremos.T1.ToString();
            baremosT2.Text = _baremos.T2.ToString();
            baremosT3.Text = _baremos.T3.ToString();
            baremosAlturaT1.Text = _baremos.AlturaT1.ToString();
            baremosAlturaT3.Text = _baremos.AlturaT3.ToString();

        }

        private void CargarDatosMetas()
        {
            meta1.Text = $"{_metas.Meta1 * 100}%";
            meta2.Text = $"{_metas.Meta2 * 100}%";

        }

        private void GenerarLibroCalidad_Load(object sender, EventArgs e)
        {
            CargarDatosBaremos();
            CargarDatosMetas();


        }

        private void addControls(UserControl uc)
        {
            // panelContainer.Controls.Clear();
            //  uc.Dock = DockStyle.Fill;
            // panelContainer.Controls.Add(uc);
            //  panel3.Controls.Add(uc);
            //  uc.BringToFront();
        }

        private void MostrarLoader()
        {
            // pictureBox1.Visible = true;
            // Aquí puedes iniciar tu tarea larga
            //  UC_Loader loader = new UC_Loader();
            // addControls(loader);
            loaderForm.StartPosition = FormStartPosition.CenterParent; // Aparece centrado respecto al formulario principal
            loaderForm.Show(this);
        }

        private void OcultarLoader()
        {
            //  pictureBox1.Visible = false;
            // Aquí puedes finalizar tu tarea larga
            // UC_Loader loader = new UC_Loader();
            // addControls(loader);
            loaderForm.Hide();
        }

        private void btnGenerarLibroFinal_Click(object sender, EventArgs e)
        {
          /*  MostrarLoader();
           
            Task.Run(() =>
            {*/
               

                if (string.IsNullOrEmpty(txtRutaCalidadDetalles.Text) || string.IsNullOrEmpty(txtRutaCalXOperarios.Text) 
              //  || string.IsNullOrEmpty(txtRutaReclamosDetalles.Text)
                )
                {
                    LibroExcelHelper.MostrarMensaje("Debe cargar todos los archivos solicitados.", true);
                }
                else
                {
                    if (double.TryParse(txtImporteCertificacion.Text, out double importeCertificacion))
                    {
                        _calidadController.CargarLibroExcel(txtRutaCalidadDetalles.Text, txtRutaCalXOperarios.Text, txtRutaReclamosDetalles.Text, importeCertificacion);
                    }
                    else
                    {
                        LibroExcelHelper.MostrarMensaje("Por favor ingrese un importe de certificación válido.", true);
                    }

                }

               
              /*  this.Invoke((MethodInvoker)delegate
                {
                    OcultarLoader();
                  
                });
            });*/


            
        }
    }
}