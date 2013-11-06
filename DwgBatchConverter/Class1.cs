using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;

namespace DwgBatchConverter
{
    public class Class1
    {
        public static bool ConvertVersion(string inputFile, DwgVersion version, out string outputFile)
        {
            outputFile = GetAndPrepareOutputFileName(inputFile);
            if (outputFile == "")
                return false;

            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(inputFile);
                try
                {
                    using (doc.LockDocument())
                        doc.Database.SaveAs(outputFile, version);
                    return true;
                }
                catch
                {
                    outputFile = "";
                    return false;
                }
                finally
                {
                    doc.CloseAndDiscard();
                }
            }
            catch
            {
                outputFile = "";
                return false;
            }
        }

        private static string GetAndPrepareOutputFileName(string inputName)
        {
            string dir = System.IO.Path.GetDirectoryName(inputName);
            string file = System.IO.Path.GetFileName(inputName);
            string outputDir = dir + @"\VersionConvert\";
            if (!System.IO.Directory.Exists(outputDir))
                System.IO.Directory.CreateDirectory(outputDir);

            string outputFile = outputDir + file;
            if (System.IO.File.Exists(outputFile))
            {
                var dlgResult = MessageBox.Show(
                    string.Format("{0} 已经存在，是否覆盖？", outputFile), 
                    "警告", 
                    MessageBoxButtons.YesNo);
                if (dlgResult != DialogResult.Yes)
                    outputFile = "";
            }
            return outputFile;
        }

        [CommandMethod("BatchConvert")]
        public static void BatchConvert()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Please Select Files to Convert\n");

            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(
                "Select AutoCAD files", "", "dwg", "", 
                Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);

            List<string> inputs = new List<string>();
            while (true)
            {
                ed.WriteMessage("Please Select Files to Convert. Select Cancel to Start Converting.\n");
                System.Windows.Forms.DialogResult result = ofd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    string[] names = ofd.GetFilenames();
                    ed.WriteMessage(string.Join("\n", names.ToArray()) + "\n");

                    inputs.AddRange(names);
                }
                else
                    break;
            }

            // AC1015 	AutoCAD 2000, AutoCAD 2000i, AutoCAD 2002
            // AC1018 	AutoCAD 2004, AutoCAD 2005, AutoCAD 2006
            // AC1021 	AutoCAD 2007, AutoCAD 2008, AutoCAD 2009
            // AC1024 	AutoCAD 2010, AutoCAD 2011, AutoCAD 2012
            // AC1027 	AutoCAD 2013
            inputs = inputs.Distinct().ToList();
            foreach (var input in inputs)
            {
                string output = "";
                if (!ConvertVersion(input, DwgVersion.AC1021, out output)) // target: AutoCAD 2007
                    ed.WriteMessage("Fail: " + input + "\n");
                else
                    ed.WriteMessage("Output: " + output + "\n");
            }

        }

    }
}
