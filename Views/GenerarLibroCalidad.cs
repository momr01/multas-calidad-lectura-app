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
using MultasLectura.Models;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.Views;
using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.Enums;

namespace MultasLectura
{
    public partial class GenerarLibroCalidad : Form
    {
        private readonly ILibroCalidadController _calidadController;
        private readonly BaremoModel _baremos = new();
        private readonly MetaModel _metas = new();

        //private Loader _loaderForm;

        public GenerarLibroCalidad()
        {
            InitializeComponent();
            ArchivoTextoHelper.VerificarExisteArchivoBaremos(_baremos);
            ArchivoTextoHelper.VerificarExisteArchivoMetas(_metas);
            _calidadController = new LibroCalidadController(_baremos!, _metas!);
           // _loaderForm = new Loader();
            ViewHelper.DragDropTextBox(txtRutaCalidadDetalles, txtRutaCalidadDetalles_DragEnter!, txtRutaCalidadDetalles_DragDrop!);
            ViewHelper.DragDropTextBox(txtRutaCalXOperarios, txtRutaCalXOperarios_DragEnter!, txtRutaCalXOperarios_DragDrop!);
            ViewHelper.DragDropTextBox(txtRutaReclamosDetalles, txtRutaReclamosDetalles_DragEnter!, txtRutaReclamosDetalles_DragDrop!);



        }

        private void btnCalidadDetalles_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtRutaCalidadDetalles);
        }

        private void btnReclamosDetalles_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtRutaReclamosDetalles);
        }

        private void btnCalXOperarios_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtRutaCalXOperarios);
        }


        private void GenerarLibroCalidad_Load(object sender, EventArgs e)
        {
            List<System.Windows.Forms.Label> labelsBaremos = new() {
                baremosT1,
                baremosT2,
                baremosT3,
                baremosAlturaT1,
                baremosAlturaT3
                };

            List<System.Windows.Forms.Label> labelsMetas = new() {
                meta1,
                meta2
                };


            ViewHelper.CargarDatosBaremos(labelsBaremos, _baremos);
          ViewHelper.CargarDatosMetas(labelsMetas, _metas);
        }

       /* private void MostrarLoader()
        {
            _loaderForm.StartPosition = FormStartPosition.CenterParent; // Aparece centrado respecto al formulario principal
            _loaderForm.Show(this);
        }

        private void OcultarLoader()
        {
            _loaderForm.Hide();
        }*/

        private void btnGenerarLibroFinal_Click(object sender, EventArgs e)
        {
            /*  MostrarLoader();

              Task.Run(() =>
              {*/

            //  MessageBox.Show(txtRutaCalidadDetalles.Lines.FirstOrDefault().ToString());

            string rutaCalDetalles = txtRutaCalidadDetalles.Lines.FirstOrDefault()!;
            string rutaCalXOperario = txtRutaCalXOperarios.Lines.FirstOrDefault()!;
            string rutaReclDetalles = txtRutaReclamosDetalles.Lines.FirstOrDefault()!;


            if (string.IsNullOrEmpty(rutaCalDetalles) || string.IsNullOrEmpty(rutaCalXOperario) || string.IsNullOrEmpty(rutaReclDetalles)
            )
            {
                LibroExcelHelper.MostrarMensaje("Debe cargar todos los archivos solicitados.", true);
            }
            else
            {
                if (double.TryParse(txtImporteCertificacion.Text, out double importeCertificacion))
                {
 
                    List<Dictionary<string, string>> rutas = new()
                    {
                       new() { [RutaArchivo.CalidadDetalles.ToString()] = rutaCalDetalles },
                        new() { [RutaArchivo.CalidadXOperario.ToString()] = rutaCalXOperario },
                        new() { [RutaArchivo.ReclamosDetalles.ToString()] = rutaReclDetalles }
                    };

                    List<Dictionary<string, string>> rutasValidas =  LibroExcelHelper.ProcesarPathArchivos(rutas, "+ Calidad_Lectura.xlsx");

                    if(rutasValidas.Count == rutas.Count + 1)
                    {
                        _calidadController.GenerarLibroCalidad(
                            LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadDetalles.ToString()),
                            LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadXOperario.ToString()),
                            LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.ReclamosDetalles.ToString()),
                            importeCertificacion,
                           LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Guardar.ToString())

                            );
                    }
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

       

        private void txtRutaCalidadDetalles_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtRutaCalidadDetalles);
        }

        private void txtRutaCalXOperarios_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtRutaCalXOperarios);
        }

        private void txtRutaReclamosDetalles_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtRutaReclamosDetalles);
        }

        private void txtRutaCalidadDetalles_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtRutaCalXOperarios_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);

        }

        private void txtRutaReclamosDetalles_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);

        }
    }
}