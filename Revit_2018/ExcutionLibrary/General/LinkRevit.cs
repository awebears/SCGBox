using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UIFrameworkServices;

namespace Revit_2018.ExcutionLibrary.General
{
    internal class LinkRevitUtils
    {
        public static void UnloadLink(Document document, RevitLinkType linkType)
        {
            if (RevitLinkType.IsLoaded(document, linkType.Id))
            {
                linkType.Unload(null);
            }
        }
        public static ImportPlacement GetImportPlacementFromToggleButton(UIApplication uiApplication, string tabName, string panelName, string radioButtonName)
        {
            RibbonPanel targetPanel = uiApplication.GetRibbonPanels(tabName).FirstOrDefault(x => x.Name == panelName);
            RadioButtonGroup targetButton = targetPanel.GetItems().FirstOrDefault(x => x.Name == radioButtonName) as RadioButtonGroup;
            ImportPlacement placement = (ImportPlacement)Enum.Parse(typeof(ImportPlacement), targetButton.Current.Name);
            return placement;
        }
        public static List<string> PathSelector(out List<string> safeNames)
        {
            List<String> paths = new List<string>();
            safeNames = null;
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Title = "SCGBox",
                Filter = "Files(*.rvt)|*.rvt",
                Multiselect = true
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                paths = openFileDialog.FileNames.ToList<string>();
                safeNames = openFileDialog.SafeFileNames.ToList<string>();
            }
            return paths;
        }

    }

    [Transaction(TransactionMode.Manual)]
    internal class LoadLinks : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            TaskDialog taskDialog = new TaskDialog("SCGBox") { TitleAutoPrefix = false };

            List<string> filePaths = LinkRevitUtils.PathSelector(out List<string> safeNames);//obtain filepaths
            if (filePaths.Count == 0)
            {
                return Result.Failed;
            }

            List<RevitLinkType> loadedRevitLinkTypes = new List<RevitLinkType>();//collect RevitLinkType loaded already among to be loaded file
            List<ElementId> newlyLoadedIds = new List<ElementId>();
            RevitLinkOptions options = new RevitLinkOptions(false);
            List<RevitLinkType> existedRevitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("载入链接");
                foreach (string filePath in filePaths)
                {
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                    try
                    {
                        newlyLoadedIds.Add(RevitLinkType.Create(doc, modelPath, options).ElementId);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException)
                    {
                        RevitLinkType exists = existedRevitLinkTypes.FirstOrDefault
                            (x => ModelPathUtils.ConvertModelPathToUserVisiblePath(x.GetExternalFileReference().GetAbsolutePath()).Equals(filePath));//a little complicate

                        loadedRevitLinkTypes.Add(exists);
                        continue;
                    }
                    catch
                    {
                        taskDialog.MainInstruction = "Error!";
                        taskDialog.MainContent = "请检查链接文件路径是否正确，是否具有读写权限等等后重试！";
                        taskDialog.Show();
                        trans.RollBack();
                        return Result.Failed;
                    }
                }
                trans.Commit();
            }

            //create RevitLinkInstances
            ImportPlacement placement = LinkRevitUtils.GetImportPlacementFromToggleButton(uiApp, "SCGBox", "批量链接", "internal_Link_Locates");
            List<RevitLinkType> excludRevitLinkTypes = new List<RevitLinkType>();
            loadedRevitLinkTypes.ForEach(x => excludRevitLinkTypes.Add(x));//exclud RevitLinkType which has created instances in current documetn among [loadedRevitLinkTypes]
            //reload if a revitLinkType is loaded already
            if (loadedRevitLinkTypes.Count != 0)
            {
                List<RevitLinkInstance> revitLinkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (RevitLinkType linkType in loadedRevitLinkTypes)
                {
                    linkType.Reload();//reload existed RevitLinkType
                    if (revitLinkInstances.Count != 0)
                    {
                        if (revitLinkInstances.FirstOrDefault(x => x.GetTypeId() == linkType.Id) != null)
                        {
                            excludRevitLinkTypes.Remove(linkType);
                        }
                    }
                }
            }
            //collect all RevitLinkTypeId need to create an instance to [newlyLoadIds]
            if (excludRevitLinkTypes.Count != 0)
            {
                excludRevitLinkTypes.ForEach(x => newlyLoadedIds.Add(x.Id));
            }
            if (newlyLoadedIds.Count != 0)//create instance
            {
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("创建链接实例");
                    foreach (ElementId id in newlyLoadedIds)
                    {
                        try
                        {
                            RevitLinkInstance.Create(doc, id, placement);
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            taskDialog.MainInstruction = "Warning!";
                            taskDialog.MainContent = "链接文件与主体模型不共享同一坐标系。" + "\n" + "对所有链接文件使用原点到原点定位?";
                            taskDialog.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
                            if (taskDialog.Show() == TaskDialogResult.Ok)
                            {
                                placement = ImportPlacement.Origin;
                                RevitLinkInstance.Create(doc, id, placement); ;//exception???
                                continue;
                            }
                            else
                            {
                                if (loadedRevitLinkTypes.Count != 0)
                                {
                                    QuickAccessToolBarService.performMultipleUndoRedoOperations(true, 1);//Revit Undo operation
                                }
                            }
                            return Result.Cancelled;
                        }
                        catch
                        {
                            taskDialog.MainInstruction = "Error!";
                            taskDialog.MainContent = "Fatal Error! Check link files carefully, and try again.";
                            taskDialog.Show();
                            trans.RollBack();
                            return Result.Failed;
                        }
                    }
                    trans.Commit();
                }
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class ReloadLinks : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            TaskDialog taskDialog = new TaskDialog("SCGBox") { TitleAutoPrefix = false };

            List<RevitLinkType> revitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
            List<ElementId> tobeCreateIds = new List<ElementId>();//collect RevitLinkTypeId needs to create RevitLinkInstance

            //reload RevitLinkType
            if (revitLinkTypes.Count != 0)
            {

                List<RevitLinkInstance> revitLinkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                foreach (RevitLinkType linkType in revitLinkTypes)
                {
                    linkType.Reload();
                    RevitLinkInstance instance = revitLinkInstances.FirstOrDefault(x => x.GetTypeId() == linkType.Id);
                    if (instance == null)//a instance gonna to be created if do not exist direived from the RevitLinkType.
                    {
                        tobeCreateIds.Add(linkType.Id);
                    }
                }
            }
            else
            {
                taskDialog.MainInstruction = "Info";
                taskDialog.MainContent = "没有已载入的链接！";
                taskDialog.Show();
                return Result.Cancelled;
            }

            //Create RevitLinkInstance from RevitLinkType which do not hava a instance in current document
            if (tobeCreateIds.Count != 0)
            {
                ImportPlacement placement = LinkRevitUtils.GetImportPlacementFromToggleButton(uiApp, "SCGBox", "批量链接", "internal_Link_Locates");
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("创建链接实例");
                    foreach (ElementId id in tobeCreateIds)
                    {
                        try
                        {
                            RevitLinkInstance.Create(doc, id, placement);
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                        {
                            taskDialog.MainInstruction = "Warning!";
                            taskDialog.MainContent = "链接文件与主体模型不共享同一坐标系。" + "\n" + "对所有链接文件使用原点到原点定位?";
                            taskDialog.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
                            if (taskDialog.Show() == TaskDialogResult.Ok)
                            {
                                placement = ImportPlacement.Origin;
                                RevitLinkInstance.Create(doc, id, placement); ;//exception???
                                continue;
                            }
                            else
                            {
                                trans.RollBack();
                                QuickAccessToolBarService.performMultipleUndoRedoOperations(true, 1);//Revit Undo operation
                                return Result.Cancelled;
                            }
                        }
                        catch (Exception)
                        {
                            taskDialog.MainInstruction = "Error!";
                            taskDialog.MainContent = "Fatal Error! Check link files carefully, and try again.";
                            taskDialog.Show();
                            trans.RollBack();
                            return Result.Failed;
                        }
                    }
                    trans.Commit();
                }
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class UnloadLinks : IExternalCommand
    {
        public List<ElementId> RevitLinkTypes { get; set; } = new List<ElementId>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //this operation does not need a transaction
            List<RevitLinkType> revitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
            string info = null;
            if (revitLinkTypes.Count != 0)
            {
                foreach (RevitLinkType revitLinkType in revitLinkTypes)
                {
                    try
                    {
                        LinkRevitUtils.UnloadLink(doc, revitLinkType);
                        RevitLinkTypes.Add(revitLinkType.Id);//collect RevitLinkType to be LinkRevitUtils
                    }
                    catch (Exception e)
                    {
                        info += e.GetType().ToString() + " : " + e.Message + "\n";
                        continue;
                    }
                }
                if (info != null)
                {
                    TaskDialog taskDialog = new TaskDialog("SCGBox")
                    { MainInstruction = "Error", TitleAutoPrefix = false, MainContent = info };
                    taskDialog.Show();
                    return Result.Failed;
                }
            }
            else
            {
                return Result.Cancelled;
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    internal class DeleteLinks : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //call UnloadLinks to unload links first.
            UnloadLinks unloadLinks = new UnloadLinks();
            if (unloadLinks.Execute(commandData, ref message, elements) == Result.Succeeded)
            {
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("删除链接");
                    foreach (ElementId id in unloadLinks.RevitLinkTypes)
                    {
                        doc.Delete(id);
                    }
                    trans.Commit();
                }
            }
            else
            {
                return Result.Cancelled;
            }
            return Result.Succeeded;
        }
    }
}
