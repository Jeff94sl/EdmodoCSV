namespace Edmodo_Csv
{
    partial class Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.Menus = new System.Windows.Forms.MenuStrip();
            this.archivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AbrirCSV = new System.Windows.Forms.ToolStripMenuItem();
            this.exportarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.acerdaDeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Abrir = new System.Windows.Forms.OpenFileDialog();
            this.Guardar = new System.Windows.Forms.SaveFileDialog();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.Menus.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.SuspendLayout();
            // 
            // Menus
            // 
            this.Menus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.archivoToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.Menus.Location = new System.Drawing.Point(0, 0);
            this.Menus.Name = "Menus";
            this.Menus.Size = new System.Drawing.Size(687, 24);
            this.Menus.TabIndex = 0;
            this.Menus.Text = "Menu";
            // 
            // archivoToolStripMenuItem
            // 
            this.archivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AbrirCSV,
            this.exportarToolStripMenuItem});
            this.archivoToolStripMenuItem.Name = "archivoToolStripMenuItem";
            this.archivoToolStripMenuItem.Size = new System.Drawing.Size(60, 20);
            this.archivoToolStripMenuItem.Text = "Archivo";
            // 
            // AbrirCSV
            // 
            this.AbrirCSV.Name = "AbrirCSV";
            this.AbrirCSV.Size = new System.Drawing.Size(122, 22);
            this.AbrirCSV.Text = "Abrir Csv";
            this.AbrirCSV.Click += new System.EventHandler(this.AbrirCSV_Click);
            // 
            // exportarToolStripMenuItem
            // 
            this.exportarToolStripMenuItem.Name = "exportarToolStripMenuItem";
            this.exportarToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
            this.exportarToolStripMenuItem.Text = "Exportar";
            this.exportarToolStripMenuItem.Click += new System.EventHandler(this.exportarToolStripMenuItem_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.acerdaDeToolStripMenuItem});
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // acerdaDeToolStripMenuItem
            // 
            this.acerdaDeToolStripMenuItem.Name = "acerdaDeToolStripMenuItem";
            this.acerdaDeToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
            this.acerdaDeToolStripMenuItem.Text = "Acerda de";
            this.acerdaDeToolStripMenuItem.Click += new System.EventHandler(this.acerdaDeToolStripMenuItem_Click);
            // 
            // Abrir
            // 
            this.Abrir.Filter = "csv|*.csv";
            // 
            // Guardar
            // 
            this.Guardar.Filter = "Excel|*.xlsx";
            // 
            // DGV
            // 
            this.DGV.AllowUserToAddRows = false;
            this.DGV.AllowUserToDeleteRows = false;
            this.DGV.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV.Location = new System.Drawing.Point(0, 24);
            this.DGV.Name = "DGV";
            this.DGV.ReadOnly = true;
            this.DGV.Size = new System.Drawing.Size(687, 345);
            this.DGV.TabIndex = 3;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(687, 369);
            this.Controls.Add(this.DGV);
            this.Controls.Add(this.Menus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.Menus;
            this.Name = "Main";
            this.Text = "EdmodoCSV";
            this.Menus.ResumeLayout(false);
            this.Menus.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip Menus;
        private System.Windows.Forms.ToolStripMenuItem archivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AbrirCSV;
        private System.Windows.Forms.OpenFileDialog Abrir;
        private System.Windows.Forms.SaveFileDialog Guardar;
        private System.Windows.Forms.ToolStripMenuItem exportarToolStripMenuItem;
        private System.Windows.Forms.DataGridView DGV;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem acerdaDeToolStripMenuItem;
    }
}

