using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Revit_2018.ExcutionLibrary.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Revit_2018.ExcutionLibrary.Datum
{

    [Transaction(TransactionMode.Manual)]
    internal class StandardizeLevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            //collect all levels in document;
            IList<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).ToElements().Cast<Level>().ToList();

            //order levelArray form down to up
            levels = levels.OrderBy(level => level.Elevation).ToList();

            //get the first floor item
            Level levelFirst = levels.Where(x => Math.Abs(x.Elevation) == levels.Min(y => Math.Abs(y.Elevation))).FirstOrDefault();

            //caculate the minimal floor index
            int index = 0;
            if (levels.First().Elevation < 0)
            {
                index -= levels.Where(level => level.Elevation < 0).ToList().Count;
            }

            LevelType levelType = doc.GetElement(levelFirst.GetValidTypes().First()) as LevelType;
            using (Transaction trans = new Transaction(doc))
            {

                trans.Start("Standarize Level");
                AddLevelType(doc, levelType);
                ElementId up = LoadFamily(doc, "标高标头_上.rfa");
                ElementId zero = LoadFamily(doc, "标高标头_正负零.rfa");
                ElementId down = LoadFamily(doc, "标高标头_下.rfa");

                //get elementId of LevelType
                ElementId level_Up_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "上标头");
                ElementId level_Zero_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "正负零标高");
                ElementId level_Down_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "下标头");

                BuiltInParameter parameter = BuiltInParameter.LEVEL_HEAD_TAG;
                doc.GetElement(level_Up_id).get_Parameter(parameter).Set(up);
                doc.GetElement(level_Zero_id).get_Parameter(parameter).Set(zero);
                doc.GetElement(level_Down_id).get_Parameter(parameter).Set(down);

                //standarlize level name and LevelType
                if (levelFirst.Elevation < 0)
                {
                    index++;
                }

                double elevation;
                double castedElevation;
                foreach (Level level in levels)
                {
                    elevation = level.Elevation;
                    castedElevation = UnitUtils.Convert(elevation, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_METERS);

                    //Match Level Type
                    if (level.Elevation < 0 & level.GetTypeId() != level_Down_id)
                    {
                        level.ChangeTypeId(level_Down_id);
                    }

                    else if (level.Elevation == 0 & level.GetTypeId() != level_Zero_id)
                    {
                        level.ChangeTypeId(level_Zero_id);
                    }
                    else if (level.Elevation > 0 & level.GetTypeId() != level_Up_id)
                    {
                        level.ChangeTypeId(level_Up_id);
                    }

                    //Rename level
                    if (index == 0)
                    {
                        index++;
                    }
                    level.Name = index.ToString() + "F " + castedElevation.ToString("0.00");
                    uiDoc.RefreshActiveView();
                    index++;
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }

        private void AddLevelType(Document doc, LevelType levelType)
        {

            using (SubTransaction subTransaction = new SubTransaction(doc))
            {
                subTransaction.Start();
                try
                {
                    levelType.Duplicate("上标头");
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException e)
                {
                    if (!e.Message.Contains("The name is already in use for this element type.")) { throw e; }
                }
                try
                {
                    levelType.Duplicate("正负零标高");
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException e)
                {
                    if (!e.Message.Contains("The name is already in use for this element type.")) { throw e; }
                }
                try
                {
                    levelType.Duplicate("下标头");
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException e)
                {
                    if (!e.Message.Contains("The name is already in use for this element type.")) { throw e; }
                }

                subTransaction.Commit();
            }
        }

        private ElementId LoadFamily(Document doc, string resourceName)
        {
            Element existFamily = new FilteredElementCollector(doc).OfClass(typeof(Family)).FirstOrDefault(x => x.Name == resourceName.TrimEnd(".rfa".ToCharArray()));
            if (existFamily != null)
            {
                Family family = existFamily as Family;
                return family.GetFamilySymbolIds().First();
            }

            string assemblyPath = this.GetType().Assembly.Location;
            string assmeblyName = this.GetType().Assembly.GetName().Name;
            string assmeblyFullname = assmeblyName + ".dll";
            string trimedPath = assemblyPath.TrimEnd(assmeblyFullname.ToCharArray());
            string filePath = Path.Combine(trimedPath, resourceName);
            string resourcePath = "Revit_2018" + ".Resources." + resourceName;
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(resourcePath);
            LoadFamilyFromResources.WriteToDisk(filePath, stream);

            try
            {
                ElementId result = null;
                using (SubTransaction subTransaction = new SubTransaction(doc))
                {
                    subTransaction.Start();
                    doc.LoadFamily(filePath, new FamilyLoadOption(), out Family family);
                    family.GetFamilySymbolIds().ToList().ForEach(x =>
                    {
                        FamilySymbol familySymbol = doc.GetElement(x) as FamilySymbol;
                        familySymbol.Activate();
                    });
                    result = family.GetFamilySymbolIds().First();
                    subTransaction.Commit();
                }
                return result;
            }
            finally
            {
                LoadFamilyFromResources.DeleteFormDisk(filePath);
            }
            //return result;
        }

        private class FamilyLoadOption : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = false;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }
    }
}
