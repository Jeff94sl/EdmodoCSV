using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Media;
using ClosedXML.Excel;

namespace Edmodo_Csv
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        private void AbrirCSV_Click(object sender, EventArgs e)
        {
            if (Abrir.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    CsvtoDatagridView(Abrir.FileName);
                    this.Text = Abrir.SafeFileName;
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.Message);
                }        
            }
        }
        private void exportarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Exportar();
        }
        public void CsvtoDatagridView(string Csv)
        {
            DataTable Tb = new DataTable();
            string[] Lineas = File.ReadAllLines(Csv);
            DataRow Row;
            if (Lineas.Length > 0)
            {
                string PrimerLinea = Lineas[0];

                string[] Etiquetas = PrimerLinea.Split(',');

                foreach (string Encabezado in Etiquetas)
                {
                    Tb.Columns.Add(new DataColumn(Encabezado));
                }
                Tb.Columns.Add("NF");
                for (int I = 1; I < Lineas.Length; I++)
                {
                    string[] Data = Lineas[I].Split(',');
                    int ColumnIndex = 0;
                    Row = Tb.NewRow();
                    string Pattern = @"[0-9 / 0-9]";
                    Regex Re = new Regex(Pattern);
                    List<int> Su = new List<int>();
                    foreach (string Datah in Etiquetas)
                    {
                        int num = 0;
                        if (Re.IsMatch(Data[ColumnIndex]))
                        {
                            Re = new Regex("[0-9]");

                            if (Re.IsMatch(Data[ColumnIndex]))
                            {
                                num = int.Parse(Data[ColumnIndex++].Replace("\"", "").Replace("=", "").Substring(0, 2));
                                Row[Datah] = num;
                                Su.Add(num);
                            }
                            else
                            {
                                Row[Datah] = Data[ColumnIndex++].Replace("\"", "").Replace("=", "");
                            }
                        }
                        else
                        {
                            Row[Datah] = Data[ColumnIndex++].Replace("\"", "").Replace("=", "");
                        }
                    }
                    
                    Row["NF"] = Su.Sum();
                    Tb.Rows.Add(Row);
                }
                DGV.DataSource = Tb;
            }
        }
        public void Exportar()
        {
            try
            {
                DataTable Tb = new DataTable();
                foreach (DataGridViewColumn Col in DGV.Columns)
                {
                    Tb.Columns.Add(Col.HeaderText, Col.ValueType);
                }

                foreach (DataGridViewRow row in DGV.Rows)
                {
                    Tb.Rows.Add();
                    foreach (DataGridViewCell Cells in row.Cells)
                    {
                        Tb.Rows[Tb.Rows.Count - 1][Cells.ColumnIndex] = Cells.Value.ToString();
                    }
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(Tb, "Reporte");

                    wb.Worksheet(1).Cells("A1:C1").Style.Fill.BackgroundColor = XLColor.DarkGreen;
                    for (int i = 1; i <= Tb.Rows.Count; i++)
                    {

                        string cellRange = string.Format("A{0}:C{0}", i + 1);
                        if (i % 2 != 0)
                        {
                            wb.Worksheet(1).Cells(cellRange).Style.Fill.BackgroundColor = XLColor.GreenYellow;
                        }
                        else
                        {
                            wb.Worksheet(1).Cells(cellRange).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }

                    }
                    wb.Worksheet(1).Columns().AdjustToContents();
                    string file = Abrir.SafeFileName.Substring(0,Abrir.SafeFileName.Length - 4);
                    Guardar.FileName = file;
                    if (Guardar.ShowDialog() == DialogResult.OK)
                    {
                        wb.SaveAs(Guardar.FileName);
                    }
                }
                SystemSounds.Exclamation.Play();
                MessageBox.Show("Exportado Completo!!!");
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }

        private void acerdaDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Acerca Ac = new Acerca();
            Ac.ShowDialog();
        }

        private void DGV_DragDrop(object sender, DragEventArgs e)
        {
            string[] A = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string i in A)
            {
                CsvtoDatagridView(i);
            }
            //CsvtoDatagridView();
        }

        private void DGV_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }
    }
}
