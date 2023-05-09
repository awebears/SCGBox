using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Branch.Tools;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Branch
{
    [Transaction(TransactionMode.Manual)]
    internal class LinkManipulator
    {
        public List<ElementId> LoadRevitLink(Document doc, out List<RevitLinkType> tobeReloadLinkTypes)
        {
            //获取文件路径
            string filter = "Files(*.rvt)|*.rvt";
            List<string> paths = WindowsFileDialog.MultiSelect(filter: filter);
            if (paths == null)
            {
                tobeReloadLinkTypes = null;
                return null;
            }

            List<RevitLinkType> revitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();//已载入的 revitlinktype
            List<ElementId> loadLinkTypes = new List<ElementId>();
            List<RevitLinkType> tobeReloadTypes = new List<RevitLinkType>();
            using (SubTransaction subTransaction = new SubTransaction(doc))
            {
                subTransaction.Start();

                //加载RevitLinkType
                RevitLinkOptions revitLinkOperations = new RevitLinkOptions(false, new WorksetConfiguration(WorksetConfigurationOption.OpenLastViewed));
                foreach (string path in paths)
                {
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                    try
                    {
                        LinkLoadResult linkLoadResult = RevitLinkType.Create(doc, modelPath, revitLinkOperations);
                        if (linkLoadResult.LoadResult != LinkLoadResultType.LinkLoaded) continue;//成功载入
                        loadLinkTypes.Add(linkLoadResult.ElementId);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException e)
                    {
                        if (e.Message.Contains("document already contains a linked model at path path."))
                        {
                            revitLinkTypes.ForEach(x =>
                            {
                                if (!x.IsNestedLink && (ModelPathUtils.ConvertModelPathToUserVisiblePath(x.GetExternalFileReference().GetAbsolutePath()).Equals(modelPath))) tobeReloadTypes.Add(x);
                            });
                        }
                        MessageBox.Show($"{e.Message}\n{e.Source}");
                        continue;
                    }
                    catch
                    {
                        TaskDialog taskDialog = new TaskDialog("SCGBox")
                        {
                            TitleAutoPrefix = false,
                            MainInstruction = "链接失败！",
                            MainContent = $"无法链接文件：\n{path}\n请检查该文件后重试。"
                        };
                        taskDialog.Show();
                        subTransaction.RollBack();
                        tobeReloadLinkTypes = null;
                        return null;
                    }
                }
                subTransaction.Commit();
            }
            tobeReloadLinkTypes = tobeReloadTypes;
            return loadLinkTypes;
        }
        /// <summary>
        /// 重载Revit链接
        /// </summary>
        /// <param name="uiDoc"></param>
        /// <param name="NoInstances">返回已载入但没有创建 revitlinkinstance 的 revitlinktype ；仅当 reloadAll 为 true 时，该值有可能不为 null</param>
        /// <param name="reloadAll">重载所有指示标志，false 为只重载有 revitlinkinstance 的revit链接</param>
        /// <returns></returns>
        public List<RevitLinkInstance> ReloadRevitLink(Document doc, List<ElementId> preSelectedIds, out List<RevitLinkType> NoInstances, bool reloadAll = false)
        {
            #region 预先选择
            if (preSelectedIds != null && preSelectedIds.Count > 0)
            {
                List<RevitLinkInstance> revitLinkInstances = new List<RevitLinkInstance>();
                foreach (ElementId id in preSelectedIds)
                {
                    if (doc.GetElement(id) is RevitLinkInstance revitLinkInstance)
                    {
                        RevitLinkType revitLinkType = doc.GetElement(revitLinkInstance.GetTypeId()) as RevitLinkType;
                        revitLinkType.Reload();
                        revitLinkInstances.Add(revitLinkInstance);
                    }
                }
                NoInstances = null;
                return revitLinkInstances;
            }
            #endregion

            #region 未选择
            List<RevitLinkType> revitLinkTypes = new List<RevitLinkType>();
            List<RevitLinkType> linkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
            if (linkTypes.Count == 0)//无已载入链接
            {
                NoInstances = null;
                return null;
            }

            List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

            if (reloadAll)
            {
                linkTypes.ForEach(x =>
                {
                    if (!x.IsNestedLink)
                    {
                        x.Reload();
                        RevitLinkInstance revitLinkInstance = linkInstances.FirstOrDefault(y => y.GetTypeId() == x.Id);
                        if (revitLinkInstance == null)
                        {
                            revitLinkTypes.Add(x);
                        }
                    }
                });
                NoInstances = revitLinkTypes;
            }
            else
            {
                linkInstances.ForEach(x =>
                {
                    RevitLinkType revitLinkType = doc.GetElement(x.GetTypeId()) as RevitLinkType;
                    revitLinkType.Reload();
                });
                NoInstances = null;
            }
            return linkInstances;
            #endregion
        }

        public void UnLoadRevitLink(Document doc, List<ElementId> preSelectedIds, bool unloadAll = false)
        {
            #region 预先选择
            if (preSelectedIds != null && preSelectedIds.Count > 0)
            {
                foreach (ElementId id in preSelectedIds)
                {
                    if (doc.GetElement(id) is RevitLinkInstance revitLinkInstance)
                    {
                        RevitLinkType revitLinkType = doc.GetElement(revitLinkInstance.GetTypeId()) as RevitLinkType;
                        revitLinkType.Unload(null);
                    }
                }
                return;
            }
            #endregion

            #region 未选择
            if (unloadAll)
            {
                List<RevitLinkType> revitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
                revitLinkTypes.ForEach(x =>
                {
                    if (!x.IsNestedLink)
                    {
                        x.Unload(null);
                    }
                });
            }
            else
            {
                List<RevitLinkInstance> revitLinkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                revitLinkInstances.ForEach(x =>
                {
                    RevitLinkType revitLinkType = doc.GetElement(x.GetTypeId()) as RevitLinkType;
                    revitLinkType.Unload(null);
                });
            }
            #endregion
        }

        public void DeleteRevitLink(Document doc, List<ElementId> preSelectedIds, bool deleteAll = false)
        {
            #region 预先选择
            if (preSelectedIds != null && preSelectedIds.Count > 0)
            {
                using (SubTransaction subTransaction = new SubTransaction(doc))
                {
                    subTransaction.Start();
                    foreach (ElementId elementId in preSelectedIds)
                    {
                        if (doc.GetElement(elementId) is RevitLinkInstance revitLinkInstance)
                        {
                            doc.Delete(revitLinkInstance.GetTypeId());
                        }
                    }
                    subTransaction.Commit();
                }
                return;
            }
            #endregion

            #region 未选择
            if (deleteAll)
            {
                List<RevitLinkType> revitLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().ToList();
                using (SubTransaction subTransaction = new SubTransaction(doc))
                {
                    subTransaction.Start();
                    revitLinkTypes.ForEach(x =>
                    {
                        try
                        {
                            if (!x.IsNestedLink)
                            {
                                doc.Delete(x.Id);
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.InvalidObjectException) { }//先于嵌套链接删除了其父链接，再调用 x.IsNestedLink 时会触发该异常
                    });
                    subTransaction.Commit();
                }
            }
            else
            {
                List<RevitLinkInstance> revitLinkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
                using (SubTransaction subTransaction = new SubTransaction(doc))
                {
                    subTransaction.Start();
                    revitLinkInstances.ForEach(x =>
                    {
                        doc.Delete(x.GetTypeId());
                    });
                    subTransaction.Commit();
                }
            }
            #endregion
        }

        public void CreateRevitLinkInstance(Document doc, ElementId revitLinkTypeId, ImportPlacement importPlacement)
        {
            using (SubTransaction subTransaction = new SubTransaction(doc))
            {
                subTransaction.Start();
                RevitLinkInstance.Create(doc, revitLinkTypeId, importPlacement);
                subTransaction.Commit();
            }
        }
        public void CreateRevitLinkInstance(Document doc, List<ElementId> revitLinkTypeIds, ImportPlacement importPlacement)
        {
            using (SubTransaction subTransaction = new SubTransaction(doc))
            {
                subTransaction.Start();
                revitLinkTypeIds.ForEach(x => RevitLinkInstance.Create(doc, x, importPlacement));
                subTransaction.Commit();
            }
        }
    }
    internal class LoadLink : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            LinkWindow linkWindow = LinkWindow.Instance;
            linkWindow.Show();
            int index = linkWindow.importPlacementEnumIndex;
            bool reloadAll = linkWindow.manipulateAll;
            linkWindow.Close();
            return Result.Succeeded;
        }
    }
}
