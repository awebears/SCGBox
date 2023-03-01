using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UIFrameworkServices;

namespace Revit_2018.ExcutionLibrary.AC
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CorrectJoinRelationships : IExternalCommand
    {
        //private List<ElementId> canNotOpterate = new List<ElementId>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            TaskDialog warningDialog = new TaskDialog("Warning!")
            {
                MainContent = "此操作根据项目文件大小，执行时间约在5分钟到1个小时，建议在空闲时间进行此操作。\n点击【确定】继续执行\n点击【取消】取消本次操作",
                MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                DefaultButton = TaskDialogResult.Cancel
            };
            if (warningDialog.Show() != TaskDialogResult.Ok)
            {
                return Result.Cancelled;
            }

            ClassifyElements allElements = new ClassifyElements(doc);
            //IList<Element> picked = uiDoc.Selection.PickElementsByRectangle(new SelectionFilter(), "pick");

            //TakeAnction(doc, allElements.Beams, picked, "Test", false);
            #region 切换结构墙
            IList<Element> elements_TobeCutted = UnionList(allElements.Columns_Str, allElements.Beams);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Floors_Str);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Walls_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Columns_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Floors_Arc);

            TakeAnction(doc, allElements.Walls_Str, elements_TobeCutted, "Join StructureWalls", false);
            #endregion

            #region 切换结构柱
            elements_TobeCutted.Clear();
            elements_TobeCutted = UnionList(allElements.Beams, allElements.Floors_Str);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Walls_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Columns_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Floors_Arc);

            TakeAnction(doc, allElements.Columns_Str, elements_TobeCutted, "Join StructureColumns", true);
            #endregion

            #region 切换梁
            elements_TobeCutted.Clear();
            elements_TobeCutted = UnionList(allElements.Floors_Str, allElements.Walls_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Columns_Arc);
            elements_TobeCutted = UnionList(elements_TobeCutted, allElements.Floors_Arc);

            TakeAnction(doc, allElements.Beams, elements_TobeCutted, "Join Beams", false);
            #endregion

            #region simple debug
            //if (canNotOpterate.Count > 0)
            //{
            //    string path_id = @"D:\debug.csv";
            //    OutPutToCSV.ExportToCSV(canNotOpterate, path_id);

            //    string path_category = @"D:\debug_additonal.csv";
            //    List<string> categories = new List<string>();
            //    canNotOpterate.ForEach(x => categories.Add(doc.GetElement(x).Category.Name));
            //    OutPutToCSV.ExportToCSV(categories, path_category);
            //    //string info = "";
            //    //canNotOpterate.ForEach(x => info += x.ToString() + "\n");
            //    //TaskDialog.Show("Debug", info);
            //}
            #endregion
            return Result.Succeeded;
        }
        private IList<Element> UnionList(IList<Element> list_1, IList<Element> list_2)
        {
            if (list_1 != null)
            {
                if (list_2 != null)
                {
                    return list_1.Union(list_2).ToList();
                }
                return list_1;
            }
            else if (list_2 != null)
            {
                return list_2.ToList();
            }
            return new List<Element>();

        }
        private void TakeAnction(Document document, IList<Element> elements_Cutting, IList<Element> elements_Cutted, string transactionName, bool isColumn)
        {
            if (elements_Cutting.Count > 0)
            {
                using (Transaction transaction = new Transaction(document, transactionName))
                {
                    int count = elements_Cutting.Count;

                    ConnectionFailureHandler connectionFailureHandler = new ConnectionFailureHandler();
                    FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();
                    failureHandlingOptions.SetFailuresPreprocessor(connectionFailureHandler);
                    failureHandlingOptions.SetClearAfterRollback(true);
                    transaction.SetFailureHandlingOptions(failureHandlingOptions);

                    transaction.Start();
                    #region 同类图元
                    if (count > 1)
                    {
                        Element element_Cutting;
                        IList<Element> toBeCutted;
                        IList<ElementId> elementIds_Cutted;
                        IList<Element> intersectedElements;
                        for (int i = 0; i < count - 2; i++)
                        {
                            element_Cutting = elements_Cutting[i];
                            toBeCutted = elements_Cutting.ToList().GetRange(i + 1, count - 1 - i);
                            elementIds_Cutted = ToElementIds(toBeCutted);

                            intersectedElements = new FilteredElementCollector(document, elementIds_Cutted)
                            .WherePasses(GetBoundingBoxIntersectsFilter(element_Cutting, document.ActiveView)).ToList();

                            if (isColumn)
                            {
                                CutFailureReason cutFailureReason;
                                foreach (Element element_Cutted in intersectedElements)
                                    if (SolidSolidCutUtils.CanElementCutElement(element_Cutting, element_Cutted, out cutFailureReason))
                                    {
                                        SolidSolidCutUtils.AddCutBetweenSolids(document, element_Cutting, element_Cutted);
                                    }
                            }
                            else
                            {
                                intersectedElements.ToList().ForEach(x => JoinAndCut(document, element_Cutting, x, true));
                            }
                        }
                    }
                    #endregion
                    #region 不同类图元
                    if (elements_Cutted.Count > 0)
                    {
                        foreach (Element element in elements_Cutting)
                        {
                            IList<ElementId> elementIds_Cutted = ToElementIds(elements_Cutted);

                            IList<Element> intersectedElements = new FilteredElementCollector(document, elementIds_Cutted)
                                .WherePasses(GetBoundingBoxIntersectsFilter(element, document.ActiveView)).ToList();
                            intersectedElements.ToList().ForEach(x => JoinAndCut(document, x, element, true));
                        }
                    }
                    #endregion
                    document.Regenerate();
                    transaction.Commit();
                }
            }
            return;
        }
        private void JoinAndCut(Document document, Element first, Element second, bool identitcalElement)
        {
            if (identitcalElement)
            {
                try
                {
                    if (!JoinGeometryUtils.AreElementsJoined(document, first, second))
                    {
                        JoinGeometryUtils.JoinGeometry(document, first, second);
                    }
                    if (JoinGeometryUtils.IsCuttingElementInJoin(document, first, second))
                    {
                        JoinGeometryUtils.SwitchJoinOrder(document, first, second);
                    }
                }
                catch
                {
                    //canNotOpterate.Add(first.Id);
                    //canNotOpterate.Add(second.Id);
                    return;
                }
                return;
            }
            if (!JoinGeometryUtils.AreElementsJoined(document, first, second))
            {
                JoinGeometryUtils.JoinGeometry(document, first, second);

            }
            if (JoinGeometryUtils.IsCuttingElementInJoin(document, first, second))
            {
                JoinGeometryUtils.SwitchJoinOrder(document, first, second);
            }
        }
        private BoundingBoxIntersectsFilter GetBoundingBoxIntersectsFilter(Element element, View view)
        {
            XYZ boundingBox_Min = element.get_BoundingBox(view).Min;
            XYZ boundingBox_Max = element.get_BoundingBox(view).Max;

            Outline outline = new Outline(boundingBox_Min, boundingBox_Max);
            return new BoundingBoxIntersectsFilter(outline);
        }
        private IList<ElementId> ToElementIds(IList<Element> elements)
        {
            List<ElementId> ids = new List<ElementId>();
            foreach (Element element in elements)
            {
                ids.Add(element.Id);
            }
            return ids;
        }

        private class SelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category.Id == new ElementId(BuiltInCategory.OST_Floors)) return true;
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }

    internal class ConnectionFailureHandler : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failureMessageAccessors = failuresAccessor.GetFailureMessages();

            if (failureMessageAccessors.Count > 0)
            {
                failureMessageAccessors.ToList().ForEach(x => x.SetCurrentResolutionType(FailureResolutionType.DetachElements));
                failuresAccessor.ResolveFailures(failureMessageAccessors);
                return FailureProcessingResult.ProceedWithCommit;
            }
            return FailureProcessingResult.Continue;
        }
    }

    internal class ClassifyElements
    {
        #region 属性
        public IList<Element> Walls_Str { get; }
        public IList<Element> Columns_Str { get; }
        public IList<Element> Beams { get; }
        public IList<Element> Floors_Str { get; }
        public IList<Element> Walls_Arc { get; }
        public IList<Element> Columns_Arc { get; }
        public IList<Element> Floors_Arc { get; }
        #endregion
        public ClassifyElements(Document doc)
        {
            #region 过滤结构墙
            BuiltInParameter parameter_Wall_Str = BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT;
            Walls_Str = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Where(x => x.get_Parameter(parameter_Wall_Str).AsInteger() == 1).ToList();
            #endregion

            #region 过滤结构柱
            Columns_Str = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤梁
            Beams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤结构板
            BuiltInParameter paremeter_Floor_Str = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
            Floors_Str = new FilteredElementCollector(doc).OfClass(typeof(Floor)).Where(x => x.get_Parameter(paremeter_Floor_Str).AsInteger() == 1).ToList();
            #endregion

            #region 过滤建筑墙
            BuiltInParameter parameter_Wall_Arc = BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT;
            Walls_Arc = new FilteredElementCollector(doc).OfClass(typeof(Wall))
                .Where(x => x.get_Parameter(parameter_Wall_Arc).AsInteger() == 0 & (x as Wall).WallType.Kind != WallKind.Curtain).ToList();
            #endregion

            #region 过滤建筑柱
            Columns_Arc = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Columns).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤建筑板
            BuiltInParameter paremeter_Floor_Arc = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
            Floors_Arc = new FilteredElementCollector(doc).OfClass(typeof(Floor)).Where(x => x.get_Parameter(paremeter_Floor_Arc).AsInteger() == 0).ToList();
            #endregion
        }
        public ClassifyElements(Document doc, IList<ElementId> elementIds)
        {
            #region 过滤结构墙
            BuiltInParameter parameter_Wall_Str = BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT;
            Walls_Str = new FilteredElementCollector(doc, elementIds).OfClass(typeof(Wall)).Where(x => x.get_Parameter(parameter_Wall_Str).AsInteger() == 1).ToList();
            #endregion

            #region 过滤结构柱
            Columns_Str = new FilteredElementCollector(doc, elementIds).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤梁
            Beams = new FilteredElementCollector(doc, elementIds).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤结构板
            BuiltInParameter paremeter_Floor_Str = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
            Floors_Str = new FilteredElementCollector(doc, elementIds).OfClass(typeof(Floor)).Where(x => x.get_Parameter(paremeter_Floor_Str).AsInteger() == 1).ToList();
            #endregion

            #region 过滤建筑墙
            BuiltInParameter parameter_Wall_Arc = BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT;
            Walls_Arc = new FilteredElementCollector(doc, elementIds).OfClass(typeof(Wall))
                .Where(x => x.get_Parameter(parameter_Wall_Arc).AsInteger() == 0 & (x as Wall).WallType.Kind != WallKind.Curtain).ToList();
            #endregion

            #region 过滤建筑柱
            Columns_Arc = new FilteredElementCollector(doc, elementIds).OfCategory(BuiltInCategory.OST_Columns).OfClass(typeof(FamilyInstance)).ToList();
            #endregion

            #region 过滤建筑板
            BuiltInParameter paremeter_Floor_Arc = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
            Floors_Arc = new FilteredElementCollector(doc, elementIds).OfClass(typeof(Floor)).Where(x => x.get_Parameter(paremeter_Floor_Arc).AsInteger() == 0).ToList();
            #endregion
        }
    }
}