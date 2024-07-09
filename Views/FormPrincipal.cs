using Microsoft.Office.Interop.Excel;
using MultasLectura.Views.Controllers;
using MultasLectura.Views.Interfaces;
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
    public partial class FormPrincipal : Form
    {
        private readonly IFormPrincipalController _controller;
        private Dictionary<string, Form> formulariosAbiertos = new Dictionary<string, Form>();

        public FormPrincipal()
        {
            InitializeComponent();
            _controller = new FormPrincipalController();
        }


        /*    private void AbrirFormularioSecundario(string nombreFormulario)
            {
                if (formulariosAbiertos.ContainsKey(nombreFormulario))
                {
                    Form formularioExistente = formulariosAbiertos[nombreFormulario];
                    if (!formularioExistente.Visible)
                    {
                        formularioExistente.Visible = true;
                    }
                    formularioExistente.Focus();
                }
                else
                {
                    Form formularioNuevo = null;

                    switch (nombreFormulario)
                    {
                        case "GenerarLibroCalidad":
                            formularioNuevo = new GenerarLibroCalidad();
                            break;
                        default:
                            break;
                    }

                    if (formularioNuevo != null)
                    {
                        formularioNuevo.MdiParent = this;
                        formulariosAbiertos.Add(nombreFormulario, formularioNuevo);
                        formularioNuevo.FormClosed += (sender, e) => formulariosAbiertos.Remove(nombreFormulario);
                        formularioNuevo.Show();
                    }
                }
            }*/


        private void generarArchivoCalidadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // AbrirFormularioSecundario("GenerarLibroCalidad");
            _controller.AbrirFormularioSecundario("GenerarLibroCalidad", formulariosAbiertos, this);
        }

        private void generarArchivoPlazosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
           _controller.AbrirFormularioSecundario("GenerarLibroPlazos", formulariosAbiertos, this);
        }
    }
}
