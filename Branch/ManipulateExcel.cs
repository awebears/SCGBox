using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using MOIE = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Branch.Tools;

namespace Branch
{
    [Transaction(TransactionMode.Manual)]
    internal class ManipulateExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //string filePath = WindowsFileDialog.Select(filter: "Excel Files (*.xlsx,*.xls) |*.xlsx;*.xls");
            //if (filePath is null)
            //{
            //    return Result.Cancelled;
            //}
            //OpenXmlHandler openXmlHandler = new OpenXmlHandler(filePath);
            //openXmlHandler.ReOrderSheet();
            //List<List<string>> strings = openXmlHandler.ReadSheet(1);
            //openXmlHandler.AddSheet(strings);
            //openXmlHandler.Dispose();

            //string info = "";
            //strings.ForEach(x =>
            //{
            //    string tem = "";
            //    x.ForEach(y => tem += $"{y} , ");
            //    info += $"{tem}\n";
            //});
            //MessageBox.Show(info);
            //openXmlHandler.TestAdd();

            MessageBox.Show(OpenXmlHandler.IterationLetter(26));

            return Result.Succeeded;
        }
    }
}
