using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Revit_2018.ExcutionLibrary.MEP
{
    [Transaction(TransactionMode.Manual)]
    internal class CutAllMepByLine : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            ElementId current_View_Id = doc.ActiveView.Id;

            //获取文档中所有水管、风管、桥架
            IList<Pipe> pipes = new FilteredElementCollector(doc, current_View_Id).OfClass(typeof(Pipe)).ToElements().Cast<Pipe>().ToList();
            IList<Duct> ducts = new FilteredElementCollector(doc, current_View_Id).OfClass(typeof(Duct)).ToElements().Cast<Duct>().ToList();
            IList<CableTray> cableTraies = new FilteredElementCollector(doc, current_View_Id).OfClass(typeof(CableTray)).ToElements().Cast<CableTray>().ToList();

            //选择打断线
            Reference refer;
            try
            {
                refer = uiDoc.Selection.PickObject(ObjectType.Element, new SelectionFilter_ModelLine(), "选择一条用于打断的线");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            if (refer == null)
            {
                return Result.Cancelled;
            }

            ModelLine refer_Line = doc.GetElement(refer.ElementId) as ModelLine;
            LocationCurve refer_Line_Curve = refer_Line.Location as LocationCurve;
            XYZ refer_Point_S = refer_Line_Curve.Curve.GetEndPoint(0); XYZ refer_Point_E = refer_Line_Curve.Curve.GetEndPoint(1);

            //投影参照线
            XYZ point_S = new XYZ(refer_Point_S.X, refer_Point_S.Y, 0); XYZ point_E = new XYZ(refer_Point_E.X, refer_Point_E.Y, 0);
            Line line = Line.CreateBound(point_S, point_E);

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("打断管线");
                #region 打断水管
                foreach (Pipe pipe in pipes)
                {
                    LocationCurve locationCurve = pipe.Location as LocationCurve;

                    XYZ LocationCurve_S = locationCurve.Curve.GetEndPoint(0); XYZ LocationCurve_E = locationCurve.Curve.GetEndPoint(1);

                    //投影线
                    Line LocationCurve_Project;
                    try
                    {
                        XYZ LocationCurve_P_S = new XYZ(LocationCurve_S.X, LocationCurve_S.Y, 0); XYZ LocationCurve_P_E = new XYZ(LocationCurve_E.X, LocationCurve_E.Y, 0);
                        LocationCurve_Project = Line.CreateBound(LocationCurve_P_S, LocationCurve_P_E);
                    }
                    catch
                    {
                        continue;
                    }

                    //求交点
                    IntersectionResultArray intersectionResultArray = new IntersectionResultArray();
                    SetComparisonResult comparisonResult = LocationCurve_Project.Intersect(line, out intersectionResultArray);

                    XYZ intersection_Point; //交点

                    if (comparisonResult == SetComparisonResult.Disjoint)
                    {
                        continue;
                    }
                    if (intersectionResultArray.IsEmpty)
                    {
                        continue;
                    }
                    intersection_Point = intersectionResultArray.get_Item(0).XYZPoint;
                    //打断
                    XYZ intereupt_Point = locationCurve.Curve.Project(intersection_Point).XYZPoint;
                    try
                    {
                        InterruptElement(doc,pipe as MEPCurve, intereupt_Point);
                    }
                    catch
                    {
                        continue;
                    }
                }
                #endregion

                #region 打断风管
                foreach (Duct duct in ducts)
                {
                    LocationCurve locationCurve = duct.Location as LocationCurve;

                    XYZ LocationCurve_S = locationCurve.Curve.GetEndPoint(0); XYZ LocationCurve_E = locationCurve.Curve.GetEndPoint(1);

                    //投影线
                    Line LocationCurve_Project;
                    try
                    {
                        XYZ LocationCurve_P_S = new XYZ(LocationCurve_S.X, LocationCurve_S.Y, 0); XYZ LocationCurve_P_E = new XYZ(LocationCurve_E.X, LocationCurve_E.Y, 0);
                        LocationCurve_Project = Line.CreateBound(LocationCurve_P_S, LocationCurve_P_E);
                    }
                    catch
                    {
                        continue;
                    }

                    //求交点
                    IntersectionResultArray intersectionResultArray = new IntersectionResultArray();
                    SetComparisonResult comparisonResult = LocationCurve_Project.Intersect(line, out intersectionResultArray);

                    XYZ intersection_Point; //交点

                    if (comparisonResult == SetComparisonResult.Disjoint)
                    {
                        continue;
                    }
                    if (intersectionResultArray.IsEmpty)
                    {
                        continue;
                    }
                    intersection_Point = intersectionResultArray.get_Item(0).XYZPoint;
                    //打断
                    XYZ intereupt_Point = locationCurve.Curve.Project(intersection_Point).XYZPoint;
                    try
                    {
                        InterruptElement(doc, duct as MEPCurve, intereupt_Point);
                    }
                    catch
                    {
                        continue;
                    }
                }
                #endregion

                #region 打断桥架
                foreach (CableTray cableTray in cableTraies)
                {
                    LocationCurve locationCurve = cableTray.Location as LocationCurve;

                    XYZ LocationCurve_S = locationCurve.Curve.GetEndPoint(0); XYZ LocationCurve_E = locationCurve.Curve.GetEndPoint(1);

                    //投影线
                    Line LocationCurve_Project;
                    try
                    {
                        XYZ LocationCurve_P_S = new XYZ(LocationCurve_S.X, LocationCurve_S.Y, 0); XYZ LocationCurve_P_E = new XYZ(LocationCurve_E.X, LocationCurve_E.Y, 0);
                        LocationCurve_Project = Line.CreateBound(LocationCurve_P_S, LocationCurve_P_E);
                    }
                    catch
                    {
                        continue;
                    }

                    //求交点
                    IntersectionResultArray intersectionResultArray = new IntersectionResultArray();
                    SetComparisonResult comparisonResult = LocationCurve_Project.Intersect(line, out intersectionResultArray);

                    XYZ intersection_Point; //交点

                    if (comparisonResult == SetComparisonResult.Disjoint)
                    {
                        continue;
                    }
                    if (intersectionResultArray.IsEmpty)
                    {
                        continue;
                    }
                    intersection_Point = intersectionResultArray.get_Item(0).XYZPoint;
                    //打断
                    XYZ intereupt_Point = locationCurve.Curve.Project(intersection_Point).XYZPoint;
                    try
                    {
                        InterruptElement(doc, cableTray as MEPCurve, intereupt_Point);
                    }
                    catch
                    {
                        continue;
                    }
                }
                #endregion
                trans.Commit();
            }

            return Result.Succeeded;
        }
        private class SelectionFilter_ModelLine : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category.Id == new ElementId(BuiltInCategory.OST_Lines))
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        IList<MEPCurve> InterruptElement(Document document, MEPCurve mepCurve, XYZ point)
        {
            IList<MEPCurve> mepCurveList = new List<MEPCurve>();

            LocationCurve locationCurve = mepCurve.Location as LocationCurve;
            XYZ start_point = locationCurve.Curve.GetEndPoint(0);
            XYZ end_point = locationCurve.Curve.GetEndPoint(1);
            Curve line1 = Line.CreateBound(start_point, point);
            Curve line2 = Line.CreateBound(point, end_point);

            //原地复制原机电管线
            MEPCurve newMEPCurve = document.GetElement(ElementTransformUtils.CopyElement(document, mepCurve.Id, new XYZ(0, 0, 0)).ElementAt(0)) as MEPCurve;
            LocationCurve newLocationCurve = newMEPCurve.Location as LocationCurve;

            //为复制的机电管线指定基线
            newLocationCurve.Curve = line1;
            //断开原机电管线的连接，建立新机电管线的连接
            ConnectorSetIterator connectorSetIterator = mepCurve.ConnectorManager.Connectors.ForwardIterator();
            while (connectorSetIterator.MoveNext())
            {
                Connector connector = connectorSetIterator.Current as Connector;
                if (connector.Origin.IsAlmostEqualTo(start_point))
                {
                    ConnectorSet tempConnectorSet = connector.AllRefs;
                    ConnectorSetIterator tempConnectorSetIterator = tempConnectorSet.ForwardIterator();
                    while (tempConnectorSetIterator.MoveNext())
                    {
                        Connector tempConnector = tempConnectorSetIterator.Current as Connector;
                        if (tempConnector != null)
                        {
                            if (tempConnector.ConnectorType == ConnectorType.End ||
                            tempConnector.ConnectorType == ConnectorType.Curve ||
                            tempConnector.ConnectorType == ConnectorType.Physical)
                            {
                                if (tempConnector.Owner.UniqueId != mepCurve.UniqueId)
                                {
                                    connector.DisconnectFrom(tempConnector);
                                    ConnectorSetIterator newConnectorSetIterator = newMEPCurve.ConnectorManager.Connectors.ForwardIterator();
                                    while (newConnectorSetIterator.MoveNext())
                                    {
                                        Connector newConnector = newConnectorSetIterator.Current as Connector;
                                        if (newConnector.Origin.IsAlmostEqualTo(start_point))
                                        {
                                            newConnector.ConnectTo(tempConnector);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //缩短原机电管线
            locationCurve.Curve = line2;
            mepCurveList.Add(newMEPCurve);
            mepCurveList.Add(mepCurve);
            return mepCurveList;
        }
    }
}
