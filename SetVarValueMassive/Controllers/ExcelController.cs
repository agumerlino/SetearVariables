using ExcelDataReader;
using SetVarValueMassive.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace SetVarValueMassive.Controllers
{
    public class ExcelController
    {
        string messageError = "";
        string messageOk = "";
        string currentController = "EXCELCONTROLLER";

        // Método privado para obtener un objeto IExcelDataReader a partir del archivo Excel
        private IExcelDataReader GetExcelReader(string filePath)
        {
            messageError = "";
            messageOk = "";
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.Read);                
                return ExcelReaderFactory.CreateReader(stream);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al obtener lector de Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(messageOk, currentController);
                return null;
            }            
        }

        // Método privado para obtener la tabla de datos a partir del archivo Excel
        private DataTable GetExcelDataTable(string filePath)
        {
            messageError = "";
            messageOk = "";
            try
            {
                using (var reader = GetExcelReader(filePath))
                {
                    
                    DataSet result = reader.AsDataSet();                    
                    return result.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al leer tabla de datos del Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(messageOk, currentController);
                return null;
            }
        }
        
        // Metodo para generar una lista de IDs del archivo Excel
        public List<int> GenerarListaDeIDExcel(string filePath)
        {
            messageError = "";
            messageOk = "";
            try
            {
                List<int> excelIDs = new List<int>();
                DataTable table = GetExcelDataTable(filePath);
                // Iterar sobre las filas de la tabla (se ignora la primera fila que contiene los nombres de las columnas)
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    int id = int.Parse(table.Rows[i][0].ToString());
                    excelIDs.Add(id);
                }
                if (excelIDs.Count == 0)
                {
                    messageError = "Error al generar lista de IDs del archivo excel";
                }
                else messageOk = $"Lista de IDs generada correctamente. Cantidad de datos: {excelIDs.Count}";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
                return excelIDs;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al generar lista de ID a partir del Excel" , MessageBoxButtons.OK, MessageBoxIcon.Error);                
                return null;
            }
        }

        // Método para obtener los nombres de las variables a partir del archivo Excel
        public List<string> ObtenerNombresVariablesExcel(string filePath)
        {
            messageError = "";
            messageOk = "";
            try
            {
                List<string> nombresVariablesExcel = new List<string>();
                DataTable table = GetExcelDataTable(filePath);

                for (int i = 0; i < table.Columns.Count; i++)
                {
                    nombresVariablesExcel.Add(table.Rows[0][i].ToString());
                }
                if (nombresVariablesExcel == null)
                {
                    messageError = "Error al obtener los nombres de las variables dentro del archivo Excel";
                }
                else messageOk = $"Nombres de variables de obtenidos correctamente. Cantidad de datos: {nombresVariablesExcel.Count}";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
                return nombresVariablesExcel;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al obtener variables del Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);        
                return null;
            }
        }

        // Método para verificar si existe una columna "ID" en el archivo Excel
        public bool VerificarColumnaID(string filePath)
        {
            messageError = "";
            messageOk = "";
            try
            {
                DataTable table = GetExcelDataTable(filePath);
                if(table.Rows[0][0].ToString().ToUpper() == "ID")
                {
                    messageOk = $"Columna ID encontrada ";                    
                    Errors.ShowMessage(messageOk, currentController);
                    return true;
                }
                else
                {
                    messageError = "Error al encontrar columna con IDs";
                    if (messageOk != "")
                    {
                        Errors.ShowMessage(messageOk, currentController);
                    }
                    if (messageError != "")
                    {
                        Errors.ShowMessage(messageError, currentController);
                    }
                    return false;
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error de verificacion de columna contenedora de IDs", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false; 
            }
        }

        // Método para obtener el número de filas en el archivo Excel
        public int ObtenerNumeroFilas(string filePath)
        {
            messageError = "";
            messageOk = "";
            try
            {
                DataTable table = GetExcelDataTable(filePath);
                if (table.Rows.Count == 0)
                {
                    messageError = "El archivo excel no contiene filas";
                }
                else messageOk = $"Numero de filas obtenido correctamente: {table.Rows.Count}";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
                return table.Rows.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al obtener el numero de filas del Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }

        // Método para obtener el valor de una celda específica en el archivo Excel
        public string ObtenerValorCelda(string filePath, string variable, int id)
        {
            messageError = "";
            messageOk = "";
            try
            {
                string cellValue = "";
                DataTable table = GetExcelDataTable(filePath);

                for (int i = 1; i < table.Rows.Count; i++)
                {
                    for (int j = 1; j < table.Columns.Count; j++)
                    {
                        if (id.ToString() == table.Rows[i][0].ToString() && variable == table.Rows[0][j].ToString())
                        {
                            cellValue = table.Rows[i][j].ToString();
                            if (cellValue != "")
                            {
                                messageOk = $"Valor de celda obtenido correctamente: {cellValue}; Variable: {variable}";
                            }
                            if (cellValue == "")
                            {
                                cellValue = " ";
                            }
                            if (messageOk != "")
                            {
                                Errors.ShowMessage(messageOk, currentController);
                            }
                            if (messageError != "")
                            {
                                Errors.ShowMessage(messageError, currentController);
                            }
                            return cellValue;
                        }
                    }
                }
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
                return cellValue;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al obtener valor de celda", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        // Método para leer todos los datos del archivo Excel y almacenarlos en una lista de listas de cadenas
        public List<List<string>> LeerExcel(string filePath)
        {
            messageError = "";
            messageOk = "";
            List<List<string>> datos = new List<List<string>>();
            try
            {
                DataTable table = GetExcelDataTable(filePath);

                foreach (DataRow row in table.Rows)
                {
                    List<string> fila = new List<string>();
                    foreach (DataColumn col in table.Columns)
                    {
                        fila.Add(row[col].ToString());// Agregar el valor de la celda a la lista de la fila
                    }
                    datos.Add(fila);// Agregar la fila a la lista de datos
                }
                if (datos.Count == 0)
                {
                    messageError = "No se pudo obtener datos del archivo Excel";
                }
                else messageOk = $"Datos del Excel obtenidos correctamente. Cantidad de filas: {datos.Count}";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
                return datos;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error al obtener datos del Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
    }    
}
