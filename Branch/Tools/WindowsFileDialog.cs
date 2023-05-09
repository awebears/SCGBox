using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Branch.Tools
{
    static class WindowsFileDialog
    {/// <summary>
     /// windows文件选择对话窗口工具
     /// </summary>
     /// <param name="title">对话窗口标题</param>
     /// <param name="filter">文件选择过滤器</param>
     /// <param name="multiselect">多选指示器</param>
     /// <returns></returns>
        public static string Select(string title = "SCGBox", string filter = "All Files(*.*)|*.*")
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Title = title,
                CheckFileExists = true,
                CheckPathExists = true,
                ReadOnlyChecked = true,
                Filter = filter,
                Multiselect = false
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            return null;
        }
        /// <summary>
        /// windows文件选择对话窗口，多选
        /// </summary>
        /// <param name="title"></param>
        /// <param name="filter"></param>
        /// <returns>用户点击取消按钮时，返回 null </returns>
        public static List<string> MultiSelect(string title = "SCGBox", string filter = "All Files(*.*)|*.*")
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Title = title,
                CheckFileExists = true,
                CheckPathExists = true,
                ReadOnlyChecked = true,
                Filter = filter,
                Multiselect = true
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileNames.ToList<string>();
            }
            return null;
        }
        public static string Save(string title = "SCGBox", string filter = "All Files(*.*)|*.*", string extension = ".xlsx")
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                Title = title,
                Filter = filter,
                DefaultExt = extension,
                AddExtension = true,
            };
            if ((saveFileDialog.ShowDialog() == DialogResult.OK))
            {
                return saveFileDialog.FileName;
            }
            return null;
        }
    }

}
