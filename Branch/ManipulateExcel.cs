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
using Parameter = Autodesk.Revit.DB.Parameter;

namespace Branch
{
    [Transaction(TransactionMode.Manual)]
    internal class ManipulateExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<Element> rebars = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_SpecialityEquipment).ToElements().ToList();

            List<List<string>> lists = new List<List<string>>() { new List<string>() { "型号", "长度" } };

            rebars.ForEach(x =>
            {
                List<string> strings = new List<string>();
                strings.Add(x.Name);
                strings.Add(x.LookupParameter("净长度").AsValueString());
                lists.Add(strings);
            });

            List<Element> links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).OfCategory(BuiltInCategory.OST_RvtLinks).ToElements().ToList();
            if (links.Count > 0)
            {
                links.ForEach(x =>
                {
                    RevitLinkInstance revitLinkInstance = x as RevitLinkInstance;
                    Document linkDoc = (x as RevitLinkInstance).GetLinkDocument();
                    if (linkDoc != null)
                    {
                        List<Element> rebars_inlink = new FilteredElementCollector(linkDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_SpecialityEquipment).ToElements().ToList();

                        rebars_inlink.ForEach(y =>
                        {
                            List<string> strings = new List<string>();
                            strings.Add(y.Name);
                            strings.Add(y.LookupParameter("净长度").AsValueString());
                            lists.Add(strings);
                        });
                    }
                });
            }

            string newFilepath = WindowsFileDialog.Save(filter: "Excel Files (*.xlsx,*.xls) |*.xlsx;*.xls");
            if (newFilepath is null)
            {
                return Result.Cancelled;
            }
            OpenXmlHandler newHandler = new OpenXmlHandler(OpenXmlHandler.CreateWorkbook(newFilepath));

            newHandler.AddSheet(lists);

            newHandler.Dispose();

            return Result.Succeeded;
        }
    }
}
