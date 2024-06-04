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
using SetVarValueMassive.Utilities;

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
        string messageError = "";
        string messageOk = "";
        string currentController = "EXCELCONTROLLER";
        public SetearVarForm()
        {
            InitializeComponent();
        }

        private void SetearVarForm_Load(object sender, EventArgs e)
        {
            btnSelectFile.Enabled = false;
            btnSetVar.Enabled = false;
            progressBarSetearVar.Visible = false;
            lblSeteoCompleted.Visible = false;
            IEdmVault7 vault = (IEdmVault7)vault1;
            if (!vault.IsLoggedIn)
            {
                vault.LoginAuto("PRUEBA", this.Handle.ToInt32());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            lblSeteoCompleted.Visible = false;
            lblEstado.Text = "";
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
                    messageOk = "Estructura del archivo Excel correcta.";
                }
                else
                {
                    lblEstado.Text = "Wrong file";
                    lblEstado.ForeColor = Color.Red;
                    btnSetVar.Enabled = false;
                }
            }
            if (messageOk != "")
            {
                Errors.ShowMessage(messageOk, currentController);
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
                messageError = $"Error. Las variables detalladas en su archivo de excel no existen o no coinciden con variables reales en el vault, elija orto archivo.{Environment.NewLine}";                
            }
            columnaID = excelController.VerificarColumnaID(filePath);
            if(columnaID == false)
            {
                messageError += "No se encontro una columna referida a los ID de los archivos.";
                MessageBox.Show(messageError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (messageOk != "")
            {
                Errors.ShowMessage(messageOk, currentController);
            }
            if (messageError != "")
            {
                Errors.ShowMessage(messageError, currentController);
            }
            bool estructuraCorrecta = columnaID && variablesCorrectas;

            return estructuraCorrecta;
        }

        private void btnSetVar_Click(object sender, EventArgs e)
        {
            lblSeteoCompleted.Visible = false;
            List<List<string>> datosExcel = excelController.LeerExcel(filePath);
            string nombreExcel = Path.GetFileName(filePath);
            VistaPreviaExcelForm form2 = new VistaPreviaExcelForm(datosExcel, numRows,nombreExcel);
            DialogResult result = form2.ShowDialog(); 

            // Verificar la decisión del usuario en el segundo formulario
            if (result == DialogResult.OK) 
            {
                progressBarSetearVar.Visible = true;
                setearVariablesController.ProgressChanged += SetearVariablesController_ProgressChanged;
                // Continuar con el proceso
                setearVariablesController.SetearVariables(idsArchivos, excelIDs, nombresVariablesExcel, fileNames, workWithFiles, filePath, this.Handle.ToInt32());
                if(progressBarSetearVar.Value == 100)
                {
                    lblSeteoCompleted.ForeColor = Color.Green;
                    progressBarSetearVar.Visible = false;
                    lblSeteoCompleted.Visible = true;
                    btnSetVar.Enabled = false;
                    btnSelectFile.Enabled = false;
                    rdFiles.Enabled = false;    
                    rdFolder.Enabled = false;   
                }
            }
            else if (result == DialogResult.Cancel) 
            {
            }
        }

        private void SetearVariablesController_ProgressChanged(object sender, int progress)
        {
            // Actualiza el valor de la ProgressBar
            progressBarSetearVar.Value = progress;
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

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
