using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            Level levelFirst = levels.Where(x => Math.Abs(x.Elevation) == levels.Min(y => Math.Abs(y.Elevation))).LastOrDefault();

            //get elementId of LevelType
            ElementId level_Up_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "上标头");
            ElementId level_Zero_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "正负零标高");
            ElementId level_Down_id = levels.First().GetValidTypes().First(x => doc.GetElement(x).Name == "下标头");

            //caculate the minimal floor index
            int index = 0;
            if (levels.First().Elevation < 0)
            {
                index = index - levels.Where(level => level.Elevation < 0).ToList().Count;
            }

            using (Transaction trans = new Transaction(doc))
            {
                //standarlize level name and LevelType
                if (levelFirst.Elevation < 0)
                {
                    index++;
                }

                trans.Start("Standarize Level");
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
    }
}
