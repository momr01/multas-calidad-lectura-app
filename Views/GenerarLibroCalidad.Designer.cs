namespace MultasLectura
{
    partial class GenerarLibroCalidad
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txtRutaCalidadDetalles = new TextBox();
            btnCalidadDetalles = new Button();
            txtRutaReclamosDetalles = new TextBox();
            groupBox1 = new GroupBox();
            txtImporteCertificacion = new TextBox();
            label3 = new Label();
            txtRutaCalXOperarios = new TextBox();
            btnCalXOperarios = new Button();
            btnReclamosDetalles = new Button();
            btnGenerarLibroFinal = new Button();
            groupBox2 = new GroupBox();
            baremosAlturaT3 = new Label();
            baremosAlturaT1 = new Label();
            baremosT3 = new Label();
            label8 = new Label();
            label7 = new Label();
            label6 = new Label();
            baremosT2 = new Label();
            label4 = new Label();
            baremosT1 = new Label();
            label2 = new Label();
            groupBox3 = new GroupBox();
            meta2 = new Label();
            meta1 = new Label();
            label10 = new Label();
            label9 = new Label();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            SuspendLayout();
            // 
            // txtRutaCalidadDetalles
            // 
            txtRutaCalidadDetalles.Location = new Point(6, 43);
            txtRutaCalidadDetalles.Name = "txtRutaCalidadDetalles";
            txtRutaCalidadDetalles.ReadOnly = true;
            txtRutaCalidadDetalles.Size = new Size(364, 29);
            txtRutaCalidadDetalles.TabIndex = 0;
            // 
            // btnCalidadDetalles
            // 
            btnCalidadDetalles.BackColor = SystemColors.ActiveCaption;
            btnCalidadDetalles.Location = new Point(376, 25);
            btnCalidadDetalles.Name = "btnCalidadDetalles";
            btnCalidadDetalles.Size = new Size(196, 63);
            btnCalidadDetalles.TabIndex = 1;
            btnCalidadDetalles.Text = "Cargar Archivo Calidad Detalles";
            btnCalidadDetalles.UseVisualStyleBackColor = false;
            btnCalidadDetalles.Click += btnCalidadDetalles_Click;
            // 
            // txtRutaReclamosDetalles
            // 
            txtRutaReclamosDetalles.Location = new Point(6, 112);
            txtRutaReclamosDetalles.Name = "txtRutaReclamosDetalles";
            txtRutaReclamosDetalles.ReadOnly = true;
            txtRutaReclamosDetalles.Size = new Size(364, 29);
            txtRutaReclamosDetalles.TabIndex = 2;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(txtImporteCertificacion);
            groupBox1.Controls.Add(label3);
            groupBox1.Controls.Add(txtRutaCalXOperarios);
            groupBox1.Controls.Add(btnCalXOperarios);
            groupBox1.Controls.Add(btnReclamosDetalles);
            groupBox1.Controls.Add(btnCalidadDetalles);
            groupBox1.Controls.Add(txtRutaCalidadDetalles);
            groupBox1.Controls.Add(txtRutaReclamosDetalles);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(578, 292);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            groupBox1.Text = "Subir Archivos";
            // 
            // txtImporteCertificacion
            // 
            txtImporteCertificacion.Location = new Point(254, 238);
            txtImporteCertificacion.Name = "txtImporteCertificacion";
            txtImporteCertificacion.Size = new Size(318, 29);
            txtImporteCertificacion.TabIndex = 7;
            txtImporteCertificacion.Text = "105147160,4";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(6, 241);
            label3.Name = "label3";
            label3.Size = new Size(242, 21);
            label3.TabIndex = 6;
            label3.Text = "Importe de Certificación Lectura=";
            // 
            // txtRutaCalXOperarios
            // 
            txtRutaCalXOperarios.Location = new Point(6, 180);
            txtRutaCalXOperarios.Name = "txtRutaCalXOperarios";
            txtRutaCalXOperarios.ReadOnly = true;
            txtRutaCalXOperarios.Size = new Size(364, 29);
            txtRutaCalXOperarios.TabIndex = 5;
            // 
            // btnCalXOperarios
            // 
            btnCalXOperarios.BackColor = SystemColors.ActiveCaption;
            btnCalXOperarios.Location = new Point(376, 162);
            btnCalXOperarios.Name = "btnCalXOperarios";
            btnCalXOperarios.Size = new Size(196, 63);
            btnCalXOperarios.TabIndex = 4;
            btnCalXOperarios.Text = "Cargar Archivo Calidad por Operarios";
            btnCalXOperarios.UseVisualStyleBackColor = false;
            btnCalXOperarios.Click += btnCalXOperarios_Click;
            // 
            // btnReclamosDetalles
            // 
            btnReclamosDetalles.BackColor = SystemColors.ActiveCaption;
            btnReclamosDetalles.Location = new Point(376, 94);
            btnReclamosDetalles.Name = "btnReclamosDetalles";
            btnReclamosDetalles.Size = new Size(196, 63);
            btnReclamosDetalles.TabIndex = 3;
            btnReclamosDetalles.Text = "Cargar Archivo Reclamos Detalles";
            btnReclamosDetalles.UseVisualStyleBackColor = false;
            btnReclamosDetalles.Click += btnReclamosDetalles_Click;
            // 
            // btnGenerarLibroFinal
            // 
            btnGenerarLibroFinal.BackColor = SystemColors.HotTrack;
            btnGenerarLibroFinal.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btnGenerarLibroFinal.ForeColor = Color.White;
            btnGenerarLibroFinal.Location = new Point(12, 332);
            btnGenerarLibroFinal.Name = "btnGenerarLibroFinal";
            btnGenerarLibroFinal.Size = new Size(853, 42);
            btnGenerarLibroFinal.TabIndex = 5;
            btnGenerarLibroFinal.Text = "GENERAR ARCHIVO CALIDAD";
            btnGenerarLibroFinal.UseVisualStyleBackColor = false;
            btnGenerarLibroFinal.Click += btnGenerarLibroFinal_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(baremosAlturaT3);
            groupBox2.Controls.Add(baremosAlturaT1);
            groupBox2.Controls.Add(baremosT3);
            groupBox2.Controls.Add(label8);
            groupBox2.Controls.Add(label7);
            groupBox2.Controls.Add(label6);
            groupBox2.Controls.Add(baremosT2);
            groupBox2.Controls.Add(label4);
            groupBox2.Controls.Add(baremosT1);
            groupBox2.Controls.Add(label2);
            groupBox2.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox2.Location = new Point(596, 12);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(269, 167);
            groupBox2.TabIndex = 6;
            groupBox2.TabStop = false;
            groupBox2.Text = "Baremos";
            // 
            // baremosAlturaT3
            // 
            baremosAlturaT3.AutoSize = true;
            baremosAlturaT3.Location = new Point(109, 134);
            baremosAlturaT3.Name = "baremosAlturaT3";
            baremosAlturaT3.Size = new Size(19, 21);
            baremosAlturaT3.TabIndex = 8;
            baremosAlturaT3.Text = "0";
            // 
            // baremosAlturaT1
            // 
            baremosAlturaT1.AutoSize = true;
            baremosAlturaT1.Location = new Point(110, 105);
            baremosAlturaT1.Name = "baremosAlturaT1";
            baremosAlturaT1.Size = new Size(19, 21);
            baremosAlturaT1.TabIndex = 8;
            baremosAlturaT1.Text = "0";
            // 
            // baremosT3
            // 
            baremosT3.AutoSize = true;
            baremosT3.Location = new Point(50, 79);
            baremosT3.Name = "baremosT3";
            baremosT3.Size = new Size(19, 21);
            baremosT3.TabIndex = 7;
            baremosT3.Text = "0";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(5, 134);
            label8.Name = "label8";
            label8.Size = new Size(98, 21);
            label8.TabIndex = 6;
            label8.Text = "ALTURA T3=";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(6, 106);
            label7.Name = "label7";
            label7.Size = new Size(98, 21);
            label7.TabIndex = 5;
            label7.Text = "ALTURA T1=";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(6, 79);
            label6.Name = "label6";
            label6.Size = new Size(38, 21);
            label6.TabIndex = 4;
            label6.Text = "T3=";
            // 
            // baremosT2
            // 
            baremosT2.AutoSize = true;
            baremosT2.Location = new Point(50, 52);
            baremosT2.Name = "baremosT2";
            baremosT2.Size = new Size(19, 21);
            baremosT2.TabIndex = 3;
            baremosT2.Text = "0";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(6, 52);
            label4.Name = "label4";
            label4.Size = new Size(38, 21);
            label4.TabIndex = 2;
            label4.Text = "T2=";
            // 
            // baremosT1
            // 
            baremosT1.AutoSize = true;
            baremosT1.Location = new Point(50, 25);
            baremosT1.Name = "baremosT1";
            baremosT1.Size = new Size(19, 21);
            baremosT1.TabIndex = 1;
            baremosT1.Text = "0";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 25);
            label2.Name = "label2";
            label2.Size = new Size(38, 21);
            label2.TabIndex = 0;
            label2.Text = "T1=";
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(meta2);
            groupBox3.Controls.Add(meta1);
            groupBox3.Controls.Add(label10);
            groupBox3.Controls.Add(label9);
            groupBox3.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox3.Location = new Point(596, 210);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(269, 94);
            groupBox3.TabIndex = 7;
            groupBox3.TabStop = false;
            groupBox3.Text = "Metas";
            // 
            // meta2
            // 
            meta2.AutoSize = true;
            meta2.Location = new Point(85, 61);
            meta2.Name = "meta2";
            meta2.Size = new Size(36, 21);
            meta2.TabIndex = 9;
            meta2.Text = "0 %";
            // 
            // meta1
            // 
            meta1.AutoSize = true;
            meta1.Location = new Point(85, 31);
            meta1.Name = "meta1";
            meta1.Size = new Size(36, 21);
            meta1.TabIndex = 8;
            meta1.Text = "0 %";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(6, 62);
            label10.Name = "label10";
            label10.Size = new Size(73, 21);
            label10.TabIndex = 1;
            label10.Text = "META 2=";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(6, 31);
            label9.Name = "label9";
            label9.Size = new Size(73, 21);
            label9.TabIndex = 0;
            label9.Text = "META 1=";
            // 
            // GenerarLibroCalidad
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(877, 413);
            Controls.Add(groupBox1);
            Controls.Add(groupBox3);
            Controls.Add(groupBox2);
            Controls.Add(btnGenerarLibroFinal);
            Name = "GenerarLibroCalidad";
            Text = "Generar Libro Calidad";
            Load += GenerarLibroCalidad_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TextBox txtRutaCalidadDetalles;
        private Button btnCalidadDetalles;
        private TextBox txtRutaReclamosDetalles;
        private GroupBox groupBox1;
        private TextBox txtRutaCalXOperarios;
        private Button btnCalXOperarios;
        private Button btnReclamosDetalles;
        private Button btnGenerarLibroFinal;
        private GroupBox groupBox2;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label baremosT2;
        private Label label4;
        private Label baremosT1;
        private Label label2;
        private GroupBox groupBox3;
        private Label label10;
        private Label label9;
        private Label baremosAlturaT3;
        private Label baremosAlturaT1;
        private Label baremosT3;
        private Label meta2;
        private Label meta1;
        private TextBox txtImporteCertificacion;
        private Label label3;
    }
}