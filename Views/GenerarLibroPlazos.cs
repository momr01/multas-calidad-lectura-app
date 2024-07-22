using MultasLectura.Enums;
using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.LibroPlazos.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultasLectura.Views
{
    public partial class GenerarLibroPlazos : Form
    {
        //private readonly ILibroCalidadController _calidadController;
        private readonly ILibroPlazosController _plazosController;
        private readonly BaremoModel _baremos = new();
        private readonly MetaModel _metas = new();

        public GenerarLibroPlazos()
        {
            InitializeComponent();
            ArchivoTextoHelper.VerificarExisteArchivoBaremos(_baremos);
            ArchivoTextoHelper.VerificarExisteArchivoMetas(_metas);
            _plazosController = new LibroPlazosController(_baremos);
            // _calidadController = new LibroCalidadController(_baremos!, _metas!);
            // _loaderForm = new Loader();
            ViewHelper.DragDropTextBox(txtRutaPlazosDetalles, txtRutaPlazosDetalles_DragEnter!, txtRutaPlazosDetalles_DragDrop!);
            // ViewHelper.DragDropTextBox(txtRutaCalXOperarios, txtRutaCalXOperarios_DragEnter!, txtRutaCalXOperarios_DragDrop!);
            // ViewHelper.DragDropTextBox(txtRutaReclamosDetalles, txtRutaReclamosDetalles_DragEnter!, txtRutaReclamosDetalles_DragDrop!);
        }

        private void txtRutaPlazosDetalles_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtRutaPlazosDetalles);
        }

        private void txtRutaPlazosDetalles_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void GenerarLibroPlazos_Load(object sender, EventArgs e)
        {
            List<System.Windows.Forms.Label> labelsBaremos = new() {
                baremosT1,
                baremosT2,
                baremosT3,
                baremosAlturaT1,
                baremosAlturaT3
                };


            ViewHelper.CargarDatosBaremos(labelsBaremos, _baremos);
        }

        private void btnPlazosDetalles_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtRutaPlazosDetalles);
        }

        private void btnGenerarLibroFinal_Click(object sender, EventArgs e)
        {
            /*  MostrarLoader();

  Task.Run(() =>
  {*/

            //  MessageBox.Show(txtRutaCalidadDetalles.Lines.FirstOrDefault().ToString());

            string rutaPlazosDetalles = txtRutaPlazosDetalles.Lines.FirstOrDefault()!;
          


            if (string.IsNullOrEmpty(rutaPlazosDetalles))
            {
                LibroExcelHelper.MostrarMensaje("Debe cargar todos los archivos solicitados.", true);
            }
            else
            {
               // if (double.TryParse(txtImporteCertificacion.Text, out double importeCertificacion))
               // {

                    List<Dictionary<string, string>> rutas = new()
                    {
                       new() { [RutaArchivo.PlazosDetalle.ToString()] = rutaPlazosDetalles }
                      
                    };

                    List<Dictionary<string, string>> rutasValidas = LibroExcelHelper.ProcesarPathArchivos(rutas, "+ Plazos_Lectura.xlsx");

                    if (rutasValidas.Count == rutas.Count + 1)
                    {
                    _plazosController.GenerarLibroPlazos(
                        LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.PlazosDetalle.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Guardar.ToString())
                        );

                    /* _calidadController.GenerarLibroCalidad(
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadDetalles.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadXOperario.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.ReclamosDetalles.ToString()),
                         importeCertificacion,
                        LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Guardar.ToString())*/


                    //  );
                    MessageBox.Show("perfecttt");
                }
               // }
               // else
               // {
               //     LibroExcelHelper.MostrarMensaje("Por favor ingrese un importe de certificación válido.", true);
               // }

            }


            /*  this.Invoke((MethodInvoker)delegate
              {
                  OcultarLoader();

              });
          });*/

        }
    }
}
