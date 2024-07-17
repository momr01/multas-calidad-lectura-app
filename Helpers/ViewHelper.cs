using MultasLectura.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Helpers
{
    public class ViewHelper
    {
        static public void DragDropTextBox(TextBox txt, DragEventHandler DragEnter, DragEventHandler DragDrop)
        {
            txt.AllowDrop = true;

            txt.DragDrop += DragDrop;
            txt.DragEnter += DragEnter;

        }

        static public void AgregarRutaATextBox(DragEventArgs e, TextBox txt)
        {
            if (e.Data!.GetDataPresent(DataFormats.FileDrop))
            {

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    txt.AppendText(file + Environment.NewLine);
                }



            }
        }

        static public void EventoDragEnter(DragEventArgs e)
        {
            if (e.Data!.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        static public void CargarDatosBaremos(List<Label> label, BaremoModel baremos)
        {
            label[0].Text = $"$ {baremos.T1:N2}";
            label[1].Text = $"$ {baremos.T2:N2}";
            label[2].Text = $"$ {baremos.T3:N2}";
            label[3].Text = $"$ {baremos.AlturaT1:N2}"; ;
            label[4].Text = $"$ {baremos.AlturaT3:N2}";

        }

        static public void CargarDatosMetas(List<Label> label, MetaModel metas)
        {
            label[0].Text = $"{metas.Meta1 * 100}%";
            label[1].Text = $"{metas.Meta2 * 100}%";

        }
    }
}
