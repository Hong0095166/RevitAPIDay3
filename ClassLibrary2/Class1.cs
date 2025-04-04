using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace C0011_Parameter_Beam_Column
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector columnCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter columnFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
            IList<Element> columnList = columnCollector.WherePasses(columnFilter).WhereElementIsNotElementType().ToElements();


            StringBuilder output = new StringBuilder("Tất cả các Structural Column trong BIM model: " + columnList.Count() + "\r\n");


            foreach (Element e1 in columnList)
            {
                string elemName = "ID: " + e1.Id.ToString() + "\n";


                ElementType type = doc.GetElement(e1.GetTypeId()) as ElementType;
                if (type != null)
                {

                    Parameter h1 = type.LookupParameter("h");
                    double hh1 = h1 != null ? ChangeUnitFeetToMillimeter(h1.AsDouble()) : 0;

                    Parameter b = type.LookupParameter("b");
                    double bb1 = b != null ? ChangeUnitFeetToMillimeter(b.AsDouble()) : 0;


                    elemName += "h = " + hh1.ToString() + " mm\n";
                    elemName += "b = " + bb1.ToString() + " mm\n";
                }
                output.Append(elemName);
                output.AppendLine();
            }
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            List<double> wallThicknessList = new List<double>();

            if (selectedIds.Count > 0)
            {
                output.AppendLine("Danh sách chiều dày của tường đã chọn:");

                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is Wall wall)
                    {
                        double thickness = ChangeUnitFeetToMillimeter(wall.Width);
                        wallThicknessList.Add(thickness);
                        output.AppendLine($"Wall ID: {id.IntegerValue}, Thickness: {thickness} mm");
                    }
                }
            }
            else
            {
                output.AppendLine("Không có tường nào được chọn.");
            }
            MessageBox.Show(output.ToString());
            return Result.Succeeded;
        }
        static double ChangeUnitFeetToMillimeter(double value)
        {
            return Math.Round(UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters), 2);
        }
    }
}