﻿using Autodesk.Revit.Attributes;
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
            string filePath = Branch.Tools.WindowsFileDialog.Select("SCGBox", "Excel Files (*.xlsx,*.xls) |*.xlsx;*.xls");
            if (filePath is null)
            {
                return Result.Cancelled;
            }
            ExcelHandler excelHandler = new ExcelHandler(filePath);
            excelHandler.AddWorkSheeet("TEST02",1);
            excelHandler.Close();
            return Result.Succeeded;
        }
    }
}