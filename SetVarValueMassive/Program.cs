using System;
using System.Runtime.InteropServices;
using EPDM.Interop.epdm;
using System.Windows.Forms;

namespace SetVarValueMassive

{
    [ComVisible(true)]
    [Guid("3A601AFC-7007-46A7-9E71-D3BD41B5E2E2")]
    public class Program : IEdmAddIn5
    {
        const int cmdId = 1;
        SetearVarForm myForm;

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            try
            {
                poInfo.mbsAddInName = "Setear variable a archivos de forma masiva";
                poInfo.mbsCompany = "Disegno Soft";
                poInfo.mbsDescription = "Setear valores de variables a traves de un Excel de forma masiva.";
                poInfo.mlAddInVersion = 1;

                poInfo.mlRequiredVersionMajor = 17;
                poInfo.mlRequiredVersionMinor = 0;

                poCmdMgr.AddCmd(cmdId, "Setear variable a archivos de forma masiva", (int)EdmMenuFlags.EdmMenu_ShowInMenuBarTools);
            }
            catch (COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
                {
                    if (poCmd.mlCmdID == cmdId)
                    {

                        if (myForm == null || myForm.IsDisposed)
                        {
                            myForm = new SetearVarForm(); 
                            myForm.Show();
                            
                        }
                    }
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
