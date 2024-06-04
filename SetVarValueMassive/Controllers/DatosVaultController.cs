using EPDM.Interop.epdm;
using SetVarValueMassive.Utilities;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SetVarValueMassive.Controllers
{
    public class DatosVaultController
    {
        string messageError = "";
        string messageOk = "";
        string currentController = "VAULTCONTROLLER";
        IEdmVault5 vault1 = new EdmVault5();

        //LLENO LISTA CON LAS RUTAS DE CADA ARCHIVO EN EL VAULT
        public List<string> ObtenerRutasEnVault(List<int> excelIDs)
        {
            List<string> fileNames = new List<string>();
            string rutaAgregar;
            IEdmVault7 vault = (IEdmVault7)vault1;
            try
            {
                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("PRUEBA", 0);
                }
                IEdmSearch5 search = vault.CreateSearch();
                IEdmSearchResult5 searchResult = search.GetFirstResult();
                while (searchResult != null && excelIDs != null)
                {
                    foreach (int id in excelIDs)
                    {
                        if (searchResult.ID == id)
                        {
                            rutaAgregar = searchResult.Path;
                            fileNames.Add(rutaAgregar);
                        }
                    }
                    searchResult = search.GetNextResult();
                }
                if (fileNames.Count > 0)
                {
                    messageOk = $"Lista de rutas de archivos generada correctamente. Cantidad de datos: {fileNames.Count}";
                }
                else messageError = $"Error al generar lista de rutas de archivos. ";                   
                if(messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);                 
                }
                if(messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);                 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(ex.Message, currentController);
            }            
            return fileNames;
        }

        //LLENO LISTA CON LOS IDS DE LOS ARCHIVOS EN EL VAULT
        public List<int> ObtenerIDsArchivosEnVault()
        {
            List<int> idsArchivos = new List<int>();
            IEdmVault7 vault = (IEdmVault7)vault1;
            try
            {
                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("PRUEBA", 0);
                }
                IEdmSearch5 search = vault.CreateSearch();
                IEdmSearchResult5 searchResult = search.GetFirstResult();

                while (searchResult != null)
                {
                    idsArchivos.Add(searchResult.ID);
                    searchResult = search.GetNextResult();
                }
                if (idsArchivos.Count > 0)
                {
                    messageOk = $"Lista de IDs de archivos generada correctamente: Cantidad de datos: {idsArchivos.Count}";
                }
                else messageError = "Error al generar lista de IDs";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(ex.Message, currentController);
            }
            
            return idsArchivos;
        }

        //LLENO LISTA CON LOS NOMBRES DE VARIABLES EN EL VAULT
        public List<string> ObtenerNombresVariablesVault()
        {
            IEdmVault7 vault = (IEdmVault7)vault1;
            List<string> nombresVariablesVault = new List<string>();
            try
            {
                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("PRUEBA", 0);
                }                
                IEdmVariableMgr5 variableManager = (IEdmVariableMgr5)vault;
                IEdmPos5 pos = variableManager.GetFirstVariablePosition();

                while (!pos.IsNull)
                {
                    IEdmVariable5 variable = variableManager.GetNextVariable(pos);
                    nombresVariablesVault.Add(variable.Name);
                }
                nombresVariablesVault.Add("ID");
                if (nombresVariablesVault.Count > 0)
                {
                    messageOk = $"Lista de variables dentro del vault generada correctamente. Cantidad de datos: {nombresVariablesVault.Count}";
                }
                else messageError = "Error al generar lista de variables";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(ex.Message, currentController);
            }
            
            return nombresVariablesVault;
        }

        //OBTENGO EL ID DE LA PARENTFOLDER DE UN ARCHIVO DETERMINADO POR SU ID
        public int GetParentFolderID(int id)
        {
            IEdmVault7 vault = (IEdmVault7)vault1;
            int idParentFolder = 0;
            try
            {
                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("PRUEBA", 0);
                }               

                IEdmSearch5 search = vault.CreateSearch();
                IEdmSearchResult5 searchResult = search.GetFirstResult();

                while (searchResult != null)
                {
                    if (searchResult.ID == id)
                    {
                        idParentFolder = searchResult.ParentFolderID;
                        break;
                    }
                    searchResult = search.GetNextResult();
                }
                if (idParentFolder != 0)
                {
                    messageOk = $"ID de la parent folder del archivo {searchResult.Name} obtenido correctamente: {idParentFolder}";
                }
                else messageError = $"Error al obtener el ID de la parentfolder del archivo {searchResult.Name}";
                if (messageOk != "")
                {
                    Errors.ShowMessage(messageOk, currentController);
                }
                if (messageError != "")
                {
                    Errors.ShowMessage(messageError, currentController);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errors.ShowMessage(ex.Message, currentController);
            }            
            return idParentFolder;
        }
    }
}
