using SetVarValueMassive.Controllers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SetVarValueMassive.Forms
{
    public partial class VistaPreviaExcelForm : Form
    {        
        private List<List<string>> _datosExcel;
        private int _numRows = 0;
        public ExcelController excelController = new ExcelController();
        int alturaTabla;          
        int anchoForm; 
        int altoForm;

        public VistaPreviaExcelForm(List<List<string>> datosExcel, int numRows,string nombreExcel)
        {
            InitializeComponent();
            this.Text = nombreExcel;
            _datosExcel = datosExcel;
            _numRows = numRows;
            ConfigurarDataGridView();
            CargarDatosEnDataGridView();          
        }

        //LLENO EL DATAGRIDVIEW CON LAS FILAS DEL EXCEL
        private void CargarDatosEnDataGridView()
        {
            int tamañoFilas = 0;
            foreach (List<string> fila in _datosExcel)
            {
                dataGridView1.Rows.Add(fila.ToArray());
            }      
            if(_numRows > 20)
            {
                _numRows = 20;
            }
            tamañoFilas = ((22 * _numRows) - (_numRows * 2));
            dataGridView1.Size = new Size((dataGridView1.ColumnCount * 100), tamañoFilas);            
        }

        //AGREGO COLUMNAS AL DATAGRIDVIEW DEPENDIENDO CUANTAS COLUMNAS TIENE EL EXCEL
        private void ConfigurarDataGridView()
        {
            // Limpiar columnas existentes (si las hay)
            dataGridView1.Columns.Clear();          

            // Agregar columnas basadas en el número de columnas en los datos del Excel
            if (_datosExcel.Count > 0)
            {
                int numColumnas = _datosExcel[0].Count;
                for (int i = 0; i < numColumnas; i++)
                {
                    dataGridView1.Columns.Add("Columna" + i, "Columna " + i);
                }
            }           
        }

        private void VistaPreviaExcelForm_Load(object sender, EventArgs e)
        {
            
            alturaTabla = dataGridView1.Height;             
            anchoForm = 600;
            altoForm = alturaTabla + 200;

            if ((dataGridView1.ColumnCount * 100) > anchoForm)
            {
                anchoForm = (dataGridView1.ColumnCount * 120);
            }
            this.Size = new Size(anchoForm, altoForm);
            this.MinimumSize = new Size(600, altoForm);
            dataGridView1.Location = new Point((this.ClientSize.Width - dataGridView1.Width) / 2, (this.ClientSize.Height - dataGridView1.Height) / 2);
        }        

        private void btnContinuar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
