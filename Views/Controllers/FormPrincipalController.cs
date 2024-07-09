using MultasLectura.Views.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Views.Controllers
{
    public class FormPrincipalController : IFormPrincipalController
    {

        public void AbrirFormularioSecundario(string nombreFormulario, Dictionary<string, Form> formulariosAbiertos, Form formPadre)
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
                    case "GenerarLibroPlazos":
                        formularioNuevo = new GenerarLibroPlazos();
                        break;
                    default:
                        break;
                }

                if (formularioNuevo != null)
                {
                    formularioNuevo.MdiParent = formPadre;
                    formulariosAbiertos.Add(nombreFormulario, formularioNuevo);
                    formularioNuevo.FormClosed += (sender, e) => formulariosAbiertos.Remove(nombreFormulario);
                    formularioNuevo.Show();
                }
            }
        }
    }
}
