using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using MOIE = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.Remoting.Metadata.W3cXsd2001;

namespace Branch.Tools
{
    class ExcelHandler
    {
        private string filePath;
        private int workSheetCount;
        private MOIE.Application excel;
        private MOIE.Workbooks workBooks;
        private MOIE.Workbook workBook;
        private Sheets sheets;
        public ExcelHandler(string filePath)
        {
            excel = new MOIE.Application() { Visible = true };
            workBooks = excel.Workbooks;
            this.filePath = filePath;
            try
            {
                workBook = workBooks.Open(filePath);
            }
            catch
            {
            }
            sheets = workBook.Sheets;
            workSheetCount = sheets.Count;
        }
        /// <summary>
        /// 插入一个工作表
        /// </summary>
        /// <param name="name">工作表名称</param>
        /// <param name="index">工作表插入处索引，1 为第一个工作表</param>
        public void AddWorkSheeet(string name, int index = 1)
        {
            if (index < 1) index = 1;
            if (index > workSheetCount) index = workSheetCount;
            Worksheet worksheet = workBook.Sheets.Add(After: sheets[index] as Worksheet);
            worksheet.Name = name;
        }

        public void Close()
        {
            workBook.Close(SaveChanges:true);
            workBooks.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
