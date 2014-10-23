using System;
using System.IO;
using System.Reflection;

namespace Funtech.EpdmWordToPdf
{
    public static class WordExporter
    {
        /// <summary>
        /// Exports the source (MS Word) document to PDF using reflection (avoids need to reference Microsoft.Office.Interop.Word).
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="openAfter"></param>
        /// <param name="error"></param>
        /// <returns></returns>
        public static bool TryExportToPdf(string source, string target, bool openAfter, out string error)
        {
            error = null;
            if (source == null) throw new ArgumentNullException(source);
            if (target == null) throw new ArgumentNullException(target);
            if (!File.Exists(source)) throw new FileNotFoundException("specified source file does not exist");

            const string progId = "Word.Application";
            Type tWordApp = Type.GetTypeFromProgID(progId);
            object app = Activator.CreateInstance(tWordApp);
            object doc = null;
            object documents = null;

            // Enum int equivalents for reference below...........
            // Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone = 0
            // Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto = 0
            // Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF = 17
            // Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint = 0
            // Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument = 0
            // Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent = 0
            // Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks = 0

            int wdAlertsNone = 0;
            int wdExportFormatPDF = 17;
            int wdExportOptimizeForPrint = 0;
            int wdExportAllDocument = 0;
            int wdExportDocumentContent = 0;
            int wdExportCreateNoBookmarks = 0;

            try
            {
                tWordApp.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { 0 });
                tWordApp.InvokeMember("Visible", BindingFlags.SetProperty, null, app, new object[] { true });
                documents = tWordApp.InvokeMember("Documents", BindingFlags.GetProperty, null, app, null);
                doc = documents.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, documents, new object[] 
                { 
                    source, false, true, false, "", "", false, "", "", wdAlertsNone, ""
                });

                doc.GetType().InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, doc, new object[] 
                { 
                    target, wdExportFormatPDF, openAfter, wdExportOptimizeForPrint, wdExportAllDocument, 1, 1, 
                    wdExportDocumentContent, true, true, wdExportCreateNoBookmarks, true, true, false
                });

                doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] 
                { 
                    false, null, null
                });
            }
            finally
            {
                tWordApp.InvokeMember("Quit", BindingFlags.InvokeMethod, null, app, null);
                if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                if (documents != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(documents);
                if (app != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                documents = null;
                doc = null;
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return true;
        }
    }
}
