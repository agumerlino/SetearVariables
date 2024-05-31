using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using SetVarValueMassive.Controllers;
using SetVarValueMassive.Forms;

namespace SetVarValueMassive
{
    public partial class SetearVarForm : Form
    {
        public int numRows = 0;
        public DataSet dataSet;
        public string filePath = "";
        public List<int> excelIDs = new List<int>();
        public List<int> idsArchivos = new List<int>();
        public List<string> nombresVariablesExcel = new List<string>();
        public List<string> nombresVariablesVault = new List<string>();
        public List<string> fileNames = new List<string>();
        IEdmVault5 vault1 = new EdmVault5();
        public bool workWithFiles = false;
        public ExcelController excelController = new ExcelController();
        public DatosVaultController vaultController = new DatosVaultController();
        public SetearVariablesController setearVariablesController = new SetearVariablesController();

        public SetearVarForm()
        {
            InitializeComponent();
        }

        private void SetearVarForm_Load(object sender, EventArgs e)
        {
            btnSelectFile.Enabled = false;
            btnSetVar.Enabled = false;
            IEdmVault7 vault = (IEdmVault7)vault1;
            if (!vault.IsLoggedIn)
            {
                vault.LoginAuto("PRUEBA", this.Handle.ToInt32());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Multiselect = false;
            openFileDialog1.Filter = "Archivos de Excel|*.xls;*.xlsx";

            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;

                bool estructuraCorrecta = VerificarEstructura(filePath);

                if (estructuraCorrecta)
                {
                    lblEstado.Text = "File correct!";
                    lblEstado.ForeColor = Color.Green;
                    btnSetVar.Enabled = true;
                }
                else
                {
                    lblEstado.Text = "Wrong file";
                    lblEstado.ForeColor = Color.Red;
                    btnSetVar.Enabled = false;
                }
            }
        }
        //Verifico la estructura del archivo excel y lleno listas correspondientes 
        private bool VerificarEstructura(string filePath)
        {
            if(excelIDs != null &&
               idsArchivos != null &&
               nombresVariablesExcel != null &&
               nombresVariablesVault != null &&
               fileNames != null)
            {
                excelIDs.Clear();
                idsArchivos.Clear();
                nombresVariablesExcel.Clear();
                nombresVariablesVault.Clear();
                fileNames.Clear();
            }
            

            // Llenar las listas
            excelIDs = excelController.GenerarListaDeIDExcel(filePath);
            idsArchivos = vaultController.ObtenerIDsArchivosEnVault();
            nombresVariablesExcel = excelController.ObtenerNombresVariablesExcel(filePath);
            nombresVariablesVault = vaultController.ObtenerNombresVariablesVault();
            fileNames = vaultController.ObtenerRutasEnVault(excelIDs);
            numRows = excelController.ObtenerNumeroFilas(filePath);

            bool columnaID = false;
            bool variablesCorrectas = true;
            variablesCorrectas = nombresVariablesExcel.All(var => nombresVariablesVault.Contains(var));
            if (variablesCorrectas == false)
            {
                MessageBox.Show("Las variables detalladas en su archivo de excel no existen o no coinciden con variables reales en el vault, elija orto archivo");
            }
            columnaID = excelController.VerificarColumnaID(filePath);
            if(columnaID == false)
            {
                MessageBox.Show("No se encontro una columna referida a los ID de los archivos");
            }
            bool estructuraCorrecta = columnaID && variablesCorrectas;

            return estructuraCorrecta;
        }

        private void btnSetVar_Click(object sender, EventArgs e)
        {
            List<List<string>> datosExcel = excelController.LeerExcel(filePath);
            string nombreExcel = Path.GetFileName(filePath);
            VistaPreviaExcelForm form2 = new VistaPreviaExcelForm(datosExcel, numRows,nombreExcel);
            DialogResult result = form2.ShowDialog(); 

            // Verificar la decisión del usuario en el segundo formulario
            if (result == DialogResult.OK) 
            {
                // Continuar con el proceso
                setearVariablesController.SetearVariables(idsArchivos, excelIDs, nombresVariablesExcel, fileNames, workWithFiles, filePath, this.Handle.ToInt32());
            }
            else if (result == DialogResult.Cancel) 
            {
            }
        }            

        private void rdFiles_CheckedChanged(object sender, EventArgs e)
        {
            if (rdFiles.Checked)
            {
                workWithFiles = true;
                btnSelectFile.Enabled = true;
                btnSetVar.Enabled = true;
            }
        }

        private void rdFolder_CheckedChanged(object sender, EventArgs e)
        {
            if (rdFolder.Checked)
            {
                workWithFiles = false;
                btnSelectFile.Enabled = true;
                btnSetVar.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
