using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using MOIE = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Branch.Tools
{
    class ExcelHandler
    {
        private int workSheetCount;
        private MOIE.Application excel;
        private MOIE.Workbooks workBooks;
        private MOIE.Workbook workBook;
        public Sheets sheets { get; }
        public ExcelHandler(string filePath)
        {
            excel = new MOIE.Application() { Visible = false };
            workBooks = excel.Workbooks;
            workBook = workBooks.Open(filePath);
            sheets = workBook.Sheets;
            workSheetCount = sheets.Count;
        }
        /// <summary>
        /// 插入一个工作表
        /// </summary>
        /// <param name="name">工作表名称</param>
        /// <param name="index">在指定索引后插入工作表，工作表索引从 1 开始</param>
        public void AddWorkSheeet(string name, int index = 1)
        {
            if (index < 1) index = 1;
            if (index > workSheetCount) index = workSheetCount;
            Worksheet worksheet = workBook.Sheets.Add(After: sheets[index] as Worksheet);
            worksheet.Name = name;
        }
        /// <summary>
        /// 读取工作表
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <returns></returns>
        public List<List<string>> ReadSheet(Worksheet worksheet)
        {
            if (worksheet is null) return null;
            int r = worksheet.UsedRange.Rows.Count;
            int c = worksheet.UsedRange.Columns.Count;

            List<List<string>> lists = new List<List<string>>();
            for (int x = 0; x < r; x++)
            {
                List<string> list = new List<string>();
                for (int y = 0; y < c; y++)
                {
                    try
                    {
                        Range range = worksheet.Cells[x + 1, y + 1] as Range;
                        list.Add(range.Value.ToString());
                    }
                    catch
                    {
                        list.Add(null);
                        continue;
                    }
                }
                lists.Add(list);
            }
            return lists;
        }
        /// <summary>
        /// 写入一个工作表
        /// </summary>
        /// <param name="worksheet">待写入的工作表</param>
        /// <param name="lists">待写入工作表的数据，二维List<string></param>
        public void WriteToSheet(Worksheet worksheet, List<List<string>> lists)
        {
            if (worksheet is null || lists is null) return;//参数为空，直接返回不进行任何操作

            for (int x = 0; x < lists.Count; x++)
            {
                for (int y = 0; y < lists[x].Count; y++)
                {
                    Range range = worksheet.Cells[x + 1, y + 1] as Range;
                    range.Value = lists[x][y];
                }

            }
        }
        /// <summary>
        /// 关闭工作簿，工作空间，退出Excel应用程序，释放相关资源
        /// </summary>
        public void Close()
        {
            workBook.Close(SaveChanges: true);
            workBooks.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
