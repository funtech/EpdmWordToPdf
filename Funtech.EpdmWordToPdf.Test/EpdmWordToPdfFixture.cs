using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Funtech.EpdmWordToPdf;
using System.Diagnostics;
using System.IO;

namespace Funtech.WordToPdf.Test
{
    [TestClass]
    public class EpdmWordToPdfFixture
    {
        [TestMethod]
        public void TryExportWordToPdfUsingLateBindingTest()
        {
            string binPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string source = binPath + "\\test.docx";
            string target = binPath + "\\test.pdf";

            string error;
            if (!WordExporter.TryExportToPdf(source, target, false, out error))
            {
                Console.WriteLine(error);
            }

            Assert.IsTrue(File.Exists(target));
        }
    }
}
