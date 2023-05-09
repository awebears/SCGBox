using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Branch.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Branch
{
    [Transaction(TransactionMode.Manual)]
    internal class FamilyLoadTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            string assemblyPath = this.GetType().Assembly.Location;
            string assmeblyName = this.GetType().Assembly.GetName().Name;
            string assmeblyFullname = assmeblyName+ ".dll";
            string trimedPath = assemblyPath.TrimEnd(assmeblyFullname.ToCharArray());
            string resourceName = "标高标头_上.rfa";
            string filePath = Path.Combine(trimedPath, resourceName);
            string resourcePath = assmeblyName + ".Resources." + resourceName;
            MessageBox.Show(resourcePath);

            Stream stream = this.GetType().Assembly.GetManifestResourceStream(resourcePath);
            MessageBox.Show(filePath);
            try
            {
                LoadFamilyFromResources.WriteToDisk(filePath, stream);
            }
            finally
            {
                LoadFamilyFromResources.DeleteFormDisk(filePath);

            }
            return Result.Succeeded;
        }
    }
    class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = false;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = false;
            return true;
        }
    }
}
