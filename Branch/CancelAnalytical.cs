using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Branch
{
    [Transaction(TransactionMode.Manual)]
    public class CancelAnalytical : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            if (doc.IsFamilyDocument)
            {
                TaskDialog.Show("ERROR", "Document Environment only！");
                return Result.Failed;
            }

            string analytical = "启用分析模型";
            #region 过滤具有结构分析模型的元素：梁、柱、墙、板、基础
            //梁
            List<FamilyInstance> beams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming)
                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
            //柱
            List<FamilyInstance> columns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns)
                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
            //墙
            BuiltInParameter parameter_Wall_Str = BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT;
            List<Wall> walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>()
                .Where(x => x.get_Parameter(parameter_Wall_Str).AsInteger() == 1).ToList();
            //板
            BuiltInParameter paremeter_Floor_Str = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
            List<Floor> floors = new FilteredElementCollector(doc).OfClass(typeof(Floor)).Cast<Floor>()
                .Where(x => x.get_Parameter(paremeter_Floor_Str).AsInteger() == 1).ToList();
            //基础--板式基础除外
            List<FamilyInstance> foundtions = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
            #endregion

            List<AnalyticalModel> analyticalModels = new FilteredElementCollector(doc)
                .Where(x => x.Category.CategoryType == CategoryType.AnalyticalModel).Cast<AnalyticalModel>().ToList();
            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("取消结构分析模型");
                #region solution_1
                //beams.ForEach(x =>
                //    {
                //        Parameter parameter = x.LookupParameter(analytical);
                //        if (parameter != null) parameter.Set(0);
                //    }
                //);
                //columns.ForEach(x =>
                //    {
                //        Parameter parameter = x.LookupParameter(analytical);
                //        if (parameter != null) parameter.Set(0);
                //    }
                //);
                //walls.ForEach(x =>
                //    {
                //        Parameter parameter = x.LookupParameter(analytical);
                //        if (parameter != null) parameter.Set(0);
                //    }
                //);
                //floors.ForEach(x =>
                //    {
                //        Parameter parameter = x.LookupParameter(analytical);
                //        if (parameter != null) parameter.Set(0);
                //    }
                //);
                //foundtions.ForEach(x =>
                //    {
                //        Parameter parameter = x.LookupParameter(analytical);
                //        if (parameter != null) parameter.Set(0);
                //    }
                //);
                #endregion
                #region solution_2
                analyticalModels.ForEach(x => x.Enable(false));
                #endregion
                transaction.Commit();
            }
            return Result.Succeeded;
        }
    }
}