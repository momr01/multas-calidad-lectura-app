namespace MultasLectura.Views
{
    partial class GenerarLibroPlazos
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txtRutaPlazosDetalles = new TextBox();
            btnPlazosDetalles = new Button();
            txtRutaDiasRegl = new TextBox();
            groupBox1 = new GroupBox();
            btnDiasReglamentarios = new Button();
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
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // txtRutaPlazosDetalles
            // 
            txtRutaPlazosDetalles.AllowDrop = true;
            txtRutaPlazosDetalles.Location = new Point(6, 43);
            txtRutaPlazosDetalles.Name = "txtRutaPlazosDetalles";
            txtRutaPlazosDetalles.ReadOnly = true;
            txtRutaPlazosDetalles.Size = new Size(364, 29);
            txtRutaPlazosDetalles.TabIndex = 0;
            // 
            // btnPlazosDetalles
            // 
            btnPlazosDetalles.BackColor = SystemColors.ActiveCaption;
            btnPlazosDetalles.Location = new Point(376, 25);
            btnPlazosDetalles.Name = "btnPlazosDetalles";
            btnPlazosDetalles.Size = new Size(196, 63);
            btnPlazosDetalles.TabIndex = 1;
            btnPlazosDetalles.Text = "Cargar Archivo Plazos Detalles";
            btnPlazosDetalles.UseVisualStyleBackColor = false;
            btnPlazosDetalles.Click += btnPlazosDetalles_Click;
            // 
            // txtRutaDiasRegl
            // 
            txtRutaDiasRegl.Location = new Point(6, 112);
            txtRutaDiasRegl.Name = "txtRutaDiasRegl";
            txtRutaDiasRegl.ReadOnly = true;
            txtRutaDiasRegl.Size = new Size(364, 29);
            txtRutaDiasRegl.TabIndex = 2;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btnDiasReglamentarios);
            groupBox1.Controls.Add(btnPlazosDetalles);
            groupBox1.Controls.Add(txtRutaPlazosDetalles);
            groupBox1.Controls.Add(txtRutaDiasRegl);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(578, 167);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            groupBox1.Text = "Subir Archivos";
            // 
            // btnDiasReglamentarios
            // 
            btnDiasReglamentarios.BackColor = SystemColors.ActiveCaption;
            btnDiasReglamentarios.Location = new Point(376, 94);
            btnDiasReglamentarios.Name = "btnDiasReglamentarios";
            btnDiasReglamentarios.Size = new Size(196, 63);
            btnDiasReglamentarios.TabIndex = 3;
            btnDiasReglamentarios.Text = "Cargar Archivo Días Reglamentarios";
            btnDiasReglamentarios.UseVisualStyleBackColor = false;
            // 
            // btnGenerarLibroFinal
            // 
            btnGenerarLibroFinal.BackColor = SystemColors.HotTrack;
            btnGenerarLibroFinal.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btnGenerarLibroFinal.ForeColor = Color.White;
            btnGenerarLibroFinal.Location = new Point(12, 211);
            btnGenerarLibroFinal.Name = "btnGenerarLibroFinal";
            btnGenerarLibroFinal.Size = new Size(853, 42);
            btnGenerarLibroFinal.TabIndex = 5;
            btnGenerarLibroFinal.Text = "GENERAR ARCHIVO PLAZOS";
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
            // GenerarLibroPlazos
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(877, 297);
            Controls.Add(groupBox1);
            Controls.Add(groupBox2);
            Controls.Add(btnGenerarLibroFinal);
            Name = "GenerarLibroPlazos";
            Text = "Generar Libro Plazos";
            Load += GenerarLibroPlazos_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TextBox txtRutaPlazosDetalles;
        private Button btnPlazosDetalles;
        private TextBox txtRutaDiasRegl;
        private GroupBox groupBox1;
        private Button btnDiasReglamentarios;
        private Button btnGenerarLibroFinal;
        private GroupBox groupBox2;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label baremosT2;
        private Label label4;
        private Label baremosT1;
        private Label label2;
        private Label baremosAlturaT3;
        private Label baremosAlturaT1;
        private Label baremosT3;
    }
}