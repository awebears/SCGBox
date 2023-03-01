using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Revit_2018.ExcutionLibrary.Utils;

namespace Revit_2018.ExcutionLibrary.MEP
{
    [Transaction(TransactionMode.Manual)]
    internal class SwitchOn : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            View view = doc.ActiveView;
            IList<ElementId> filterElementIds = view.GetFilters().ToList();
            if (filterElementIds.Count == 0)
            {
                return Result.Cancelled;
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("打开所有视图过滤器");

                foreach (ElementId id in filterElementIds)
                {
                    view.SetFilterVisibility(id, true);
                }

                trans.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class SwitchOff : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            View view = doc.ActiveView;
            IList<ElementId> filterElementIds = view.GetFilters().ToList();
            if (filterElementIds.Count == 0)
            {
                return Result.Cancelled;
            }
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("关闭所有视图过滤器");

                foreach (ElementId id in filterElementIds)
                {
                    view.SetFilterVisibility(id, false);
                }

                trans.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class SwitchSelectedOff : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View view = doc.ActiveView;

            //当前视图没有应用任何视图过滤器，直接返回Cancel
            IList<ElementId> filterElementIds = view.GetFilters().ToList();
            if (filterElementIds.Count == 0)
            {
                TaskDialog.Show("过滤器", "当前视图没有应用任何过滤器！");
                return Result.Cancelled;
            }


            List<Element> selected = new List<Element>();
            List<ElementId> preSelectedIds = uiDoc.Selection.GetElementIds().ToList();
            if (preSelectedIds.Count > 0)//先选择图元，后点击命令
            {
                preSelectedIds.ForEach(x => selected.Add(doc.GetElement(x)));
            }
            else//先点击命令，后选择图元
            {
                List<Reference> references;
                string assembly = "Revit_2018.ExcutionLibrary.DependenceLibrary.Gma.System.MouseKeyHook.dll";
                LoadLibrary loadLibrary = new LoadLibrary(assembly);
                MouseAndKeyBoard mouseAndKeyBoard = new MouseAndKeyBoard();
                try
                {
                    loadLibrary.Subscribe();
                    mouseAndKeyBoard.MouseAndKeyBoard_Subscribe();
                    references = uiDoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "选择图元，按空格完成选择").ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
                finally
                {
                    mouseAndKeyBoard.MouseAndKeyBoard_Unsubscribe();
                    loadLibrary.Unsubscribe();
                }
                if (references == null)
                {
                    return Result.Cancelled;
                }
                references.ForEach(x => selected.Add(doc.GetElement(x)));
            }

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("关闭选中图元的视图过滤器");
                foreach (Element element in selected)
                {
                    foreach (ElementId filterElementId in filterElementIds)
                    {
                        ParameterFilterElement filterElement = doc.GetElement(filterElementId) as ParameterFilterElement;
                        if (filterElement != null)
                        {
                            IList<FilterRule> filterRules = filterElement.GetRules();
                            int indicator = 0;//过滤规则指示器
                            foreach (FilterRule filterRule in filterRules)
                            {
                                if (filterRule.ElementPasses(element))
                                {
                                    indicator++;
                                }
                            }

                            if (indicator == filterRules.Count & indicator != 0 & view.GetFilterVisibility(filterElementId))
                            {
                                view.SetFilterVisibility(filterElementId, false);
                            }
                        }
                        else
                        {
                            SelectionFilterElement selectionFilterElement = doc.GetElement(filterElementId) as SelectionFilterElement;
                            if (selectionFilterElement.GetElementIds().Contains(element.Id) & view.GetFilterVisibility(filterElementId))
                            {
                                view.SetFilterVisibility(filterElementId, false);
                            }
                        }
                    }
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class SwitchSelectedOn : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View view = doc.ActiveView;

            //当前视图没有应用任何视图过滤器，直接返回Cancel
            IList<ElementId> filterElementIds = view.GetFilters().ToList();
            if (filterElementIds.Count == 0)
            {
                TaskDialog.Show("过滤器", "当前视图没有应用任何过滤器！");
                return Result.Cancelled;
            }

            List<Element> selected = new List<Element>();
            List<ElementId> preSelectedIds = uiDoc.Selection.GetElementIds().ToList();
            if (preSelectedIds.Count > 0)//先选择图元，后点击命令
            {
                preSelectedIds.ForEach(x => selected.Add(doc.GetElement(x)));
            }
            else//先点击命令，后选择图元
            {
                List<Reference> references;
                string assembly = "Revit_2018.ExcutionLibrary.DependenceLibrary.Gma.System.MouseKeyHook.dll";
                LoadLibrary loadLibrary = new LoadLibrary(assembly);
                MouseAndKeyBoard mouseAndKeyBoard = new MouseAndKeyBoard();
                try
                {
                    loadLibrary.Subscribe();
                    mouseAndKeyBoard.MouseAndKeyBoard_Subscribe();
                    references = uiDoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "选择图元，按空格完成选择").ToList();

                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
                finally
                {
                    mouseAndKeyBoard.MouseAndKeyBoard_Unsubscribe();
                    loadLibrary.Unsubscribe();
                }
                if (references == null)
                {
                    return Result.Cancelled;
                }
                references.ForEach(x => selected.Add(doc.GetElement(x)));
            }

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("隔离选中图元的过滤器");
                foreach (ElementId id in filterElementIds)
                {

                    ParameterFilterElement filterElement = doc.GetElement(id) as ParameterFilterElement;
                    bool indicator = false;
                    if (filterElement != null)
                    {
                        IList<FilterRule> filterRules = filterElement.GetRules();
                        foreach (Element element in selected)
                        {
                            int inter_Indicator = 0;
                            foreach (FilterRule filterRule in filterRules)
                            {
                                if (filterRule.ElementPasses(element))
                                {
                                    inter_Indicator++;
                                }
                            }
                            if (inter_Indicator == filterRules.Count)
                            {
                                indicator = true;
                                break;
                            }
                        }
                        view.SetFilterVisibility(id, indicator);
                    }
                    else
                    {
                        SelectionFilterElement selectionFilterElement = doc.GetElement(id) as SelectionFilterElement;
                        foreach (Element element in selected)
                        {
                            if (selectionFilterElement.GetElementIds().Contains(element.Id))
                            {
                                indicator = true;
                                break;
                            }
                        }
                        view.SetFilterVisibility(id, indicator);
                    }
                }

                trans.Commit();
            }
            return Result.Succeeded;
        }
    }
}