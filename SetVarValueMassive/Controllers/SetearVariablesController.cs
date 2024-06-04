using EPDM.Interop.epdm;
using SetVarValueMassive.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SetVarValueMassive.Controllers
{
    public class SetearVariablesController
    {
        string messageError = "";
        string messageOk = "";
        string currentController = "SETVARVALUES";
        private DatosVaultController vaultController = new DatosVaultController();
        private ExcelController excelController = new ExcelController(); 
        private IEdmVault5 vault1 = new EdmVault5();
        public event EventHandler<int> ProgressChanged;
        StringBuilder mensajeErrores = new StringBuilder();
        List<string> listaErrores = new List<string>();

        //CHECKIN, SETEO Y CHECKOUT DE VARIABLES 
        public void SetearVariables(List<int> idsArchivos, List<int> excelIDs, List<string> nombresVariablesExcel, List<string> fileNames, bool workWithFiles, string filePath, int formHandle)
        {

            string cellValue;
            IEdmVault7 vault = (IEdmVault7)vault1;            
            IEdmBatchUnlock2 batchUnlocker;
            EdmSelItem[] ppoSelection = new EdmSelItem[0];
            IEdmSelectionList6 fileList = null;
            int nbrFiles = fileNames.Count;
            Array.Resize(ref ppoSelection, nbrFiles);
            IEdmFile5 archivo;
            IEdmFolder5 folder;
            IEdmFolder5 ppoRetParentFolder;
            IEdmPos5 aPos;
            EdmSelectionObject poSel;
            string str;            
            int idVar;          
            bool unlock = false;            
            
            
            try
            {
                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("PRUEBA", 0);
                }
                
                batchUnlocker = (IEdmBatchUnlock2)vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
                IEdmVariableMgr5 variableMgr = (IEdmVariableMgr5)vault;
                IEdmBatchUpdate2 batchUpdate = (IEdmBatchUpdate2)vault.CreateUtility(EdmUtility.EdmUtil_BatchUpdate);
                //Recorro la lista de IDs 
                foreach (int id in excelIDs)
                {
                    //Pregunto si el ID se encuentra en el vault
                    if (idsArchivos.Contains(id))
                    {
                        //Pregunto si estoy trabajando con archivos
                        if (workWithFiles)
                        {
                            //Obtener el archivo con el ID
                            IEdmFile7 file = (IEdmFile7)vault.GetObject(EdmObjectType.EdmObject_File, id);

                            //Obtener el ID de la parentFolder con el ID del archivo
                            int parentFolderID = vaultController.GetParentFolderID(id);
                            if (file != null)
                            {
                                //Pregunto si el archivo no fue traido
                                if (!file.IsLocked)
                                {
                                    //Traer el archivo
                                    file.LockFile(parentFolderID, formHandle, 0);
                                }
                                messageOk = $"Archivo obtenido correctamente: {file.Name}";                                
                                OnProgressChanged(10);
                            }
                            else messageError = $"Error al obtener el archivo";
                            //Envio mensajes de exito o de error a la consola
                            if (messageOk != "")
                            {
                                Errors.ShowMessage(messageOk, currentController);
                            }
                            if (messageError != "")
                            {
                                Errors.ShowMessage(messageError, currentController);
                            }
                        }
                        //Recorro la lista de variables que contiene el archivo Excel
                        foreach (string variable in nombresVariablesExcel)
                        {
                            //Me posiciono en cada columna que no sea la de IDs
                            if (variable != "ID")
                            {
                                //Obtener ID de la variable a modificar
                                idVar = variableMgr.GetVariable(variable).ID;
                                //Obtener el valor de la celda correspondiente
                                cellValue = excelController.ObtenerValorCelda(filePath, variable, id);
                                if (cellValue != null)
                                {
                                    messageOk = $"Valor a modificar obtenido correctamente: {cellValue}";
                                    Errors.ShowMessage(messageOk, currentController);                                    
                                    OnProgressChanged(35);
                                }
                                if (workWithFiles)
                                {
                                    //Seteo de variable de archivo                                
                                    batchUpdate.SetVar(id, idVar, cellValue, "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                }
                                if (!workWithFiles)
                                {
                                    //Seteo de variable de carpeta                             
                                    batchUpdate.SetFolderVar(id, idVar, cellValue, (int)EdmBatchFlags.EdmBatch_Nothing);
                                }
                            }
                        }
                    }
                    else
                    {
                        //Envio mensajes de exito o de error a la consola
                        messageError = $"Error. El id {id} no se encuentra en el vault, por lo tanto se omitirá esta fila {Environment.NewLine}";
                        if (messageOk != "")
                        {
                            Errors.ShowMessage(messageOk, currentController);
                        }
                        if (messageError != "")
                        {
                            Errors.ShowMessage(messageError, currentController);
                        }
                        listaErrores.Add(messageError);                        
                    }
                }
                foreach (string error in listaErrores)
                {
                    mensajeErrores.AppendLine(error);
                }
                if(mensajeErrores.Length > 0)
                {
                    MessageBox.Show(mensajeErrores.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }               
                OnProgressChanged(45);
                //Manejo de errores al momento de actualizar variables
                EdmBatchError2[] errors;
                int errorSize = batchUpdate.CommitUpdate(out errors, null);
                if (errorSize > 0)
                {
                    StringBuilder errorMessage = new StringBuilder();
                    foreach (var error in errors)
                    {
                        errorMessage.AppendLine($"Error: {error}");
                    }
                    MessageBox.Show(errorMessage.ToString(), "Error de actualización de variables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (!workWithFiles)
                {
                    messageOk = $"Seteo de variables de carpeta realizado correctamente";
                    
                    OnProgressChanged(100);
                    MessageBox.Show(messageOk, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (messageOk != "")
                    {
                        Errors.ShowMessage(messageOk, currentController);
                    }
                    if (messageError != "")
                    {
                        Errors.ShowMessage(messageError, currentController);
                    }
                }

                //Pregunto si estoy trabajando con archivos, ya que si son carpetas no necesito registrar    
                if (workWithFiles)
                {
                    try
                    {
                        int h = 0;
                        foreach (string fileName in fileNames)
                        {
                            archivo = vault.GetFileFromPath(fileName, out ppoRetParentFolder);
                            if (archivo != null)
                            {
                                aPos = archivo.GetFirstFolderPosition();
                                folder = archivo.GetNextFolder(aPos);
                                ppoSelection[h] = new EdmSelItem();
                                ppoSelection[h].mlDocID = archivo.ID;
                                ppoSelection[h].mlProjID = folder.ID;
                                h++;
                            }
                            OnProgressChanged(46);
                        }
                        batchUnlocker.AddSelection((EdmVault5)vault, ref ppoSelection);

                        if ((batchUnlocker != null))
                        {
                            batchUnlocker.CreateTree(formHandle, (int)EdmUnlockBuildTreeFlags.Eubtf_RefreshFileListing + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock + (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_ForceUnlock);

                            batchUnlocker.Comment = "Updates";
                            fileList = (IEdmSelectionList6)batchUnlocker.GetFileList((int)EdmUnlockFileListFlag.Euflf_GetUnlocked + (int)EdmUnlockFileListFlag.Euflf_GetUndoLocked + (int)EdmUnlockFileListFlag.Euflf_GetUnprocessed);

                            aPos = fileList.GetHeadPosition();
                            OnProgressChanged(60);
                            str = "Getting " + fileList.Count + " files: ";
                            while (!(aPos.IsNull))
                            {
                                fileList.GetNext2(aPos, out poSel);
                                str = str + $"{Environment.NewLine}" + poSel.mbsPath;
                            }                            
                            OnProgressChanged(50);
                            MessageBox.Show($"{str}", "Lista de archivos a obtener", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            batchUnlocker.UnlockFiles(formHandle, null);
                            unlock = true;
                            OnProgressChanged(75);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        messageError = $"Error al registrar archivos. {ex.Message} ";
                        Errors.ShowMessage(messageError, currentController);
                    }
                    //Pregunto si se registraron los archivos
                    if (unlock)
                    {
                        mensajeErrores.Clear();
                        listaErrores.Clear();
                        // Vuelvo a recorrer para setear variable de revision, despues de haber registrado
                        foreach (int id in excelIDs)
                        {
                            if (idsArchivos.Contains(id))
                            {
                                IEdmFile7 file = (IEdmFile7)vault.GetObject(EdmObjectType.EdmObject_File, id);
                                foreach (string variable in nombresVariablesExcel)
                                {
                                    if (variable == "Revision")
                                    {
                                        cellValue = excelController.ObtenerValorCelda(filePath, variable, id);
                                        SetearRev(file, cellValue);                                        
                                    }
                                }                                
                                OnProgressChanged(85);
                            }
                        }
                        foreach (string error in listaErrores)
                        {
                            mensajeErrores.AppendLine(error);
                        }
                        if (mensajeErrores.Length > 0)
                        {
                            MessageBox.Show(mensajeErrores.ToString(), "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }                        
                        messageOk = $"Seteo de variables de archivos realizado correctamente";                        
                        OnProgressChanged(100);
                        MessageBox.Show(messageOk, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Errors.ShowMessage(messageOk, currentController);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error" , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        //Metodo para setear revisiones
        private void SetearRev(IEdmFile7 file, object revisionValue)
        {
            IEdmVault7 vault = (IEdmVault7)vault1;
            if (!vault.IsLoggedIn)
            {
                vault.LoginAuto("PRUEBA", 0);
            }
            try
            {
                //Pregunto si el valor de revision tiene la cantidad de caracteres correcta
                if (revisionValue.ToString().Length < 4 && revisionValue.ToString().Length > 1)
                {
                    IEdmRevisionMgr2 revisionManager = (IEdmRevisionMgr2)vault.CreateUtility(EdmUtility.EdmUtil_RevisionMgr);

                    bool canIncrement = false;
                    // ID del esquema de revisión
                    int revisionSchemaId = revisionManager.GetRevisionNumberIDFromFile(file.ID, out canIncrement);

                    // Si no hay un esquema de revisión disponible para ese estado, salir de la función
                    if (revisionSchemaId == 0)
                    {
                        MessageBox.Show("El archivo " + file.Name + " no posee ningun esquema de revision segun su estado en el flujo de trabajo","Error ",MessageBoxButtons.OK, MessageBoxIcon.Error);
                        messageError = $"Error al obtener el esquema de revision del archivo {file.Name}";
                        if (messageError != "")
                        {
                            Errors.ShowMessage(messageError, currentController);
                        }
                        return;
                    }

                    // Obtener detalles sobre las revisiones disponibles para el esquema de revisión
                    EdmRevNo[] revisionNumbers;
                    revisionManager.GetRevisionNumbers(revisionSchemaId, out revisionNumbers);

                    // Obtener detalles sobre los componentes de revisión
                    EdmRevComponent2[] revisionComponents;
                    revisionManager.GetRevisionNumberComponents2(-revisionSchemaId, out revisionComponents);

                    // Arreglo para almacenar los nuevos valores de los contadores de revisión              
                    EdmRevCounter[] revisionCounters = new EdmRevCounter[revisionComponents.Count()];

                    // Asigno valores de revision a sus variables correspondientes como Nombre de componente y Contador de componente
                    revisionCounters[0].mbsComponentName = revisionComponents[1].mbsComponentName;
                    revisionCounters[0].mlCounter = (Convert.ToInt32(revisionValue.ToString().ToUpper()[0]) - Convert.ToInt32('A') + 1);

                    //Separo la parte numerica de la alfabetica en el valor de revision
                    int i = 0;
                    while (i < revisionValue.ToString().Length && !char.IsDigit(revisionValue.ToString()[i]))
                    {
                        i++;
                    }
                    revisionCounters[1].mbsComponentName = revisionComponents[0].mbsComponentName;
                    revisionCounters[1].mlCounter = int.Parse(revisionValue.ToString().Substring(i));

                    revisionManager.SetRevisionCounters(file.ID, revisionCounters);

                    // Incrementar la revisión del archivo
                    revisionManager.IncrementRevision(file.ID);

                    // Confirmar los cambios de revisión
                    EdmRevError[] revisionErrors = { };
                    revisionManager.Commit("Incremento Revision", out revisionErrors);

                    messageOk = $"La revision nueva del archivo {file.Name} es:{file.CurrentRevision}";                    
                    if (messageOk != "")
                    {
                        Errors.ShowMessage(messageOk, currentController);
                    }
                    if (messageError != "")
                    {
                        Errors.ShowMessage(messageError, currentController);
                    }
                }
                else
                {                    
                    messageError = $"Error. El valor de revision para el archivo {file.Name} no es valido, por lo tanto se omitira el seteo interno de esta variable {Environment.NewLine}";
                    listaErrores.Add(messageError);
                    
                    if (messageOk != "")
                    {
                        Errors.ShowMessage(messageOk, currentController);
                    }
                    if (messageError != "")
                    {
                        Errors.ShowMessage(messageError, currentController);
                    }
                }              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        protected virtual void OnProgressChanged(int progress)
        {
            ProgressChanged?.Invoke(this, progress);
        }
    }
}
