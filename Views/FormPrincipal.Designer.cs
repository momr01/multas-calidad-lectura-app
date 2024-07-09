namespace MultasLectura.Views
{
    partial class FormPrincipal
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            menuStrip1 = new MenuStrip();
            multasLecturaToolStripMenuItem = new ToolStripMenuItem();
            generarArchivoCalidadToolStripMenuItem = new ToolStripMenuItem();
            generarArchivoReclamosToolStripMenuItem = new ToolStripMenuItem();
            generarArchivoPlazosToolStripMenuItem = new ToolStripMenuItem();
            generarArchivoResumenToolStripMenuItem = new ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { multasLecturaToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // multasLecturaToolStripMenuItem
            // 
            multasLecturaToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { generarArchivoCalidadToolStripMenuItem, generarArchivoReclamosToolStripMenuItem, generarArchivoPlazosToolStripMenuItem, generarArchivoResumenToolStripMenuItem });
            multasLecturaToolStripMenuItem.Name = "multasLecturaToolStripMenuItem";
            multasLecturaToolStripMenuItem.Size = new Size(97, 20);
            multasLecturaToolStripMenuItem.Text = "Multas Lectura";
            // 
            // generarArchivoCalidadToolStripMenuItem
            // 
            generarArchivoCalidadToolStripMenuItem.Name = "generarArchivoCalidadToolStripMenuItem";
            generarArchivoCalidadToolStripMenuItem.Size = new Size(213, 22);
            generarArchivoCalidadToolStripMenuItem.Text = "Generar Archivo Calidad";
            generarArchivoCalidadToolStripMenuItem.Click += generarArchivoCalidadToolStripMenuItem_Click;
            // 
            // generarArchivoReclamosToolStripMenuItem
            // 
            generarArchivoReclamosToolStripMenuItem.Name = "generarArchivoReclamosToolStripMenuItem";
            generarArchivoReclamosToolStripMenuItem.Size = new Size(213, 22);
            generarArchivoReclamosToolStripMenuItem.Text = "Generar Archivo Reclamos";
            // 
            // generarArchivoPlazosToolStripMenuItem
            // 
            generarArchivoPlazosToolStripMenuItem.Name = "generarArchivoPlazosToolStripMenuItem";
            generarArchivoPlazosToolStripMenuItem.Size = new Size(213, 22);
            generarArchivoPlazosToolStripMenuItem.Text = "Generar Archivo Plazos";
            generarArchivoPlazosToolStripMenuItem.Click += generarArchivoPlazosToolStripMenuItem_Click;
            // 
            // generarArchivoResumenToolStripMenuItem
            // 
            generarArchivoResumenToolStripMenuItem.Name = "generarArchivoResumenToolStripMenuItem";
            generarArchivoResumenToolStripMenuItem.Size = new Size(213, 22);
            generarArchivoResumenToolStripMenuItem.Text = "Generar Archivo Resumen";
            // 
            // FormPrincipal
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(menuStrip1);
            IsMdiContainer = true;
            MainMenuStrip = menuStrip1;
            Name = "FormPrincipal";
            Text = "Página Principal";
            WindowState = FormWindowState.Maximized;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem multasLecturaToolStripMenuItem;
        private ToolStripMenuItem generarArchivoCalidadToolStripMenuItem;
        private ToolStripMenuItem generarArchivoReclamosToolStripMenuItem;
        private ToolStripMenuItem generarArchivoPlazosToolStripMenuItem;
        private ToolStripMenuItem generarArchivoResumenToolStripMenuItem;
    }
}