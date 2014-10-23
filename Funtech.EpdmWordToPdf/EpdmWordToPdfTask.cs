using System;
using EdmLib;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace Funtech.EpdmWordToPdf
{
    [ComVisible(true), Guid("AF92C44F-3224-4A0A-A203-31B46F0B93B3")]
    public class EpdmWordToPdfTask : EdmLib.IEdmAddIn5
    {
        private static void HandleEx(Exception ex, bool showMsg)
        {
            // TODO: Implement logging...
            Trace.WriteLine(ex.ToString());
            if (showMsg) MessageBox.Show(ex.Message);
        }

        /// <summary>
        /// Sets the status of a task to EdmTaskStat_DoneFailed and sends correct message based on argument Exception.
        /// </summary>
        /// <param name="poCmd"></param>
        /// <param name="ex"></param>
        protected void SetTaskError(EdmCmd poCmd, Exception ex)
        {
            string msg;

            IEdmTaskInstance taskInst = poCmd.mpoExtra as IEdmTaskInstance;
            if (taskInst != null)
            {
                int hres = 0;
                if (ex is COMException)
                {
                    hres = (ex as COMException).ErrorCode;
                }
                var dt = DateTime.Now.ToString();
                msg = string.Format("{0} --- {1}", dt, ex.ToString());
                taskInst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, hres, msg);
            }
        }

        private void SetAddInInfo(ref EdmAddInInfo poInfo)
        {
            var a = Assembly.GetExecutingAssembly();
            poInfo.mbsAddInName =  a.GetAttributeValue<AssemblyTitleAttribute>();
            poInfo.mbsCompany = a.GetAttributeValue<AssemblyCompanyAttribute>();
            poInfo.mbsDescription = a.GetAttributeValue<AssemblyDescriptionAttribute>();
            poInfo.mlAddInVersion = a.GetMajorVersion();
            poInfo.mlRequiredVersionMajor = 14;
            poInfo.mlRequiredVersionMinor = 4;
        }

        private VaultObject[] GetVaultObjects(Array ppoData)
        {
            var vaultObjects = new VaultObject[ppoData.Length];
            for (int i = 0; i < ppoData.Length; i++)
            {
                var itemDat = (EdmCmdData)ppoData.GetValue(i);
                var objType = (EdmObjectType)itemDat.mlLongData1;

                // From API Doc.......
                // mlObjectID1 = ID of the selected object. (See, e.g, IEdmObject5.ID.) 
                // mlObjectID2 = Parent folder ID if the selected object is a file. 
                // mbsStrData1 = Complete file system path to the object. 
                // mbsStrData2 = Configuration name if the object is a file. Can be an empty string. 
                // mlLongData1 = An EdmObjectType constant telling what kind of object this is.

                if (objType == EdmObjectType.EdmObject_File ||
                    objType == EdmObjectType.EdmObject_Folder)
                {
                    vaultObjects[i] = new VaultObject
                    {
                        ObjectType = objType,
                        Id = itemDat.mlObjectID1,
                        ParentFolderId = objType == EdmObjectType.EdmObject_File ? itemDat.mlObjectID2 : 0
                    };
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
            return vaultObjects;
        }

        private void TaskRun(IEdmVault13 vault, VaultObject[] vaultObjects)
        {
            foreach (var o in vaultObjects)
            {
                // get the file object
                var file = vault.GetObject(EdmObjectType.EdmObject_File, o.Id) as IEdmFile8;

                // cache latest version of the source file
                file.GetFileCopy(0);

                // export the file
                string source = file.GetLocalPath(o.ParentFolderId.Value);
                string target = Path.GetDirectoryName(source) + "\\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf";
                string error;
                WordExporter.TryExportToPdf(source, target, false, out error);
            }
        }

        void IEdmAddIn5.GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            try
            {
                SetAddInInfo(ref poInfo);
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
            }
            catch (Exception ex)
            {
                HandleEx(ex, true);
            }
        }

        void IEdmAddIn5.OnCmd(ref EdmCmd poCmd, ref Array ppoData)
        {
            try
            {
                switch (poCmd.meCmdType)
                {
                    case EdmCmdType.EdmCmd_TaskRun:
                        try
                        {
                            TaskRun(poCmd.mpoVault as IEdmVault13, GetVaultObjects(ppoData));
                        }
                        catch (Exception ex)
                        {
                            SetTaskError(poCmd, ex);
                        }
                        break;
                    case EdmCmdType.EdmCmd_TaskSetup:
                        // TODO: Implement task setup page...
                        break;
                    case EdmCmdType.EdmCmd_TaskSetupButton:
                        var taskProps = poCmd.mpoExtra as IEdmTaskProperties;
                        taskProps.TaskFlags = (int)(
                            EdmTaskFlag.EdmTask_SupportsInitExec |
                            EdmTaskFlag.EdmTask_SupportsChangeState);
                        taskProps.SetMenuCmds(new EdmTaskMenuCmd[] 
                        {
                            new EdmTaskMenuCmd
                            {
                                mlCmdID = 1,
                                mbsMenuString = "Convert DOC to PDF",
                                mbsStatusBarHelp = "",
                                mlEdmMenuFlags = (int)(
                                    EdmMenuFlags.EdmMenu_ContextMenuItem |
                                    EdmMenuFlags.EdmMenu_OnlyFiles | 
                                    EdmMenuFlags.EdmMenu_MustHaveSelection)
                            }
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleEx(ex, true);
            }
        }
    }
}
