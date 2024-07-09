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
        private readonly BaremoModel _baremos = new BaremoModel();
        private readonly MetaModel _metas = new MetaModel();

        private Loader _loaderForm;

        public GenerarLibroCalidad()
        {
            InitializeComponent();
            ArchivoTextoHelper.VerificarExisteArchivoBaremos(_baremos);
            ArchivoTextoHelper.VerificarExisteArchivoMetas(_metas);
            _calidadController = new LibroCalidadController(_baremos!, _metas!);
            _loaderForm = new Loader();
            DragDropTextBoxes();



        }

        private void DragDropTextBoxes()
        {
            txtRutaCalidadDetalles.AllowDrop = true;
            txtRutaCalXOperarios.AllowDrop = true;
            txtRutaReclamosDetalles.AllowDrop = true;

            txtRutaCalidadDetalles.DragEnter += txtRutaCalidadDetalles_DragEnter!;
            txtRutaCalidadDetalles.DragDrop += txtRutaCalidadDetalles_DragDrop!;
            txtRutaCalXOperarios.DragEnter += txtRutaCalXOperarios_DragEnter!;
            txtRutaCalXOperarios.DragDrop += txtRutaCalXOperarios_DragDrop!;
            txtRutaReclamosDetalles.DragEnter += txtRutaReclamosDetalles_DragEnter!;
            txtRutaReclamosDetalles.DragDrop += txtRutaReclamosDetalles_DragDrop!;
        }

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
            CargarDatosBaremos();
            CargarDatosMetas();
        }

        private void MostrarLoader()
        {
            _loaderForm.StartPosition = FormStartPosition.CenterParent; // Aparece centrado respecto al formulario principal
            _loaderForm.Show(this);
        }

        private void OcultarLoader()
        {
            _loaderForm.Hide();
        }

        private void btnGenerarLibroFinal_Click(object sender, EventArgs e)
        {
            /*  MostrarLoader();

              Task.Run(() =>
              {*/

            //  MessageBox.Show(txtRutaCalidadDetalles.Lines.FirstOrDefault().ToString());

            string rutaCalDetalles = txtRutaCalidadDetalles.Lines.FirstOrDefault();
            string rutaCalXOperario = txtRutaCalXOperarios.Lines.FirstOrDefault();
            string rutaReclDetalles = txtRutaReclamosDetalles.Lines.FirstOrDefault();


            if (string.IsNullOrEmpty(rutaCalDetalles) || string.IsNullOrEmpty(rutaCalXOperario) || string.IsNullOrEmpty(rutaReclDetalles)
            )
            {
                LibroExcelHelper.MostrarMensaje("Debe cargar todos los archivos solicitados.", true);
            }
            else
            {
                if (double.TryParse(txtImporteCertificacion.Text, out double importeCertificacion))
                {
                    _calidadController.CargarLibroExcel(rutaCalDetalles, rutaCalXOperario, rutaReclDetalles, importeCertificacion);
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
            AgregarRutaATextBox(e, txtRutaCalidadDetalles);
        }

        private void txtRutaCalXOperarios_DragDrop(object sender, DragEventArgs e)
        {
            AgregarRutaATextBox(e, txtRutaCalXOperarios);
        }

        private void txtRutaReclamosDetalles_DragDrop(object sender, DragEventArgs e)
        {
            AgregarRutaATextBox(e, txtRutaReclamosDetalles);
        }

        private void AgregarRutaATextBox(DragEventArgs e, System.Windows.Forms.TextBox txt)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    txt.AppendText(file + Environment.NewLine);
                }



            }
        }

        private void EventoDragEnter(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void txtRutaCalidadDetalles_DragEnter(object sender, DragEventArgs e)
        {
            EventoDragEnter(e);
        }

        private void txtRutaCalXOperarios_DragEnter(object sender, DragEventArgs e)
        {
            EventoDragEnter(e);

        }

        private void txtRutaReclamosDetalles_DragEnter(object sender, DragEventArgs e)
        {
            EventoDragEnter(e);

        }
    }
}