using Autodesk.AutoCAD.Customization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Microsoft.Office.Interop.Excel;
using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace MTool_V1.DLL
{
    public static class OwneRevitMathods
    {
        public static void CreateLevels(Document document, List<double> Elevations, List<string> LevelsName, DisplayUnitType displayUnit = DisplayUnitType.DUT_METERS)
        {
            using (Transaction TnsCreateLevels = new Transaction(document, "Create Levels"))
            {
                TnsCreateLevels.Start();
                var ElevationsList = Elevations.ToList();
                var LevelsNameList = LevelsName.ToList();
                if (Elevations.Count() <= LevelsName.Count())
                {
                    for (int i = 0; i < Elevations.Count(); i++)
                    {
                        Level DemoLevel = Level.Create(document, UnitUtils.ConvertToInternalUnits(ElevationsList[i], displayUnit));
                        DemoLevel.Name = LevelsNameList[i];
                    }
                }
                else if (Elevations.Count() > LevelsName.Count())
                {
                    var DefOfCount = Elevations.Count() - LevelsName.Count();
                    for (int i = 0; i < LevelsName.Count(); i++)
                    {
                        Level DemoLevel = Level.Create(document, UnitUtils.ConvertToInternalUnits(ElevationsList[i], displayUnit));
                        DemoLevel.Name = LevelsNameList[i];
                    }
                    for (int i = LevelsName.Count(); i < LevelsName.Count() + DefOfCount; i++)
                    {
                        Level DemoLevel = Level.Create(document, UnitUtils.ConvertToInternalUnits(ElevationsList[i], displayUnit));
                        DemoLevel.Name = $"Level {i}";
                    }
                }
                else
                {
                    MessageBox.Show("There is missing data");
                }
                TnsCreateLevels.Commit();
            }
        }
        public static void CreateViews(Document document, List<ElementId> LevelsIDList = null, List<string> LevelsNameList = null)
        {
            using (Transaction TnsCreateView = new Transaction(document, "Create View"))
            {
                TnsCreateView.Start();
                var LevelsID = new List<ElementId>();
                var LevelsName = new List<string>();
                if (LevelsIDList == null)
                {
                    LevelsIDList = (List<ElementId>)RevitFilters.GetAllElementsDataByCategory(document, BuiltInCategory.OST_Levels, RevitFilters.ReturnElementData.ElementID, "Levels");
                    LevelsID = LevelsIDList;
                }
                else if (LevelsName == null)
                {
                    LevelsNameList = (List<string>)RevitFilters.GetAllElementsDataByCategory(document, BuiltInCategory.OST_Levels, RevitFilters.ReturnElementData.ElementName, "Levels");
                    LevelsName = LevelsNameList;
                }
                else if (LevelsID == null && LevelsName == null)
                {
                    LevelsName = (List<string>)RevitFilters.GetAllElementsDataByCategory(document, BuiltInCategory.OST_Levels, RevitFilters.ReturnElementData.ElementName, "Levels");
                    LevelsID = (List<ElementId>)RevitFilters.GetAllElementsDataByCategory(document, BuiltInCategory.OST_Levels, RevitFilters.ReturnElementData.ElementID, "Levels");
                    LevelsID = LevelsIDList;
                    LevelsName = LevelsNameList;
                }
                else
                {
                    LevelsID = LevelsIDList;
                    LevelsName = LevelsNameList;
                }
                var ViewFamilyTypeID = RevitFilters.GetViewFamilyType(document, RevitFilters.ReturnViewData.ViewFamilyTpeID);
                if (LevelsName.Count() <= LevelsID.Count())
                {
                    for (int i = 0; i < LevelsName.Count(); i++)
                    {
                        var View = ViewPlan.Create(document, ViewFamilyTypeID as ElementId, LevelsID[i] as ElementId);
                        View.Name = LevelsName[i];
                    }
                    var DeferanceBetweenTwoLists = LevelsID.Count() - LevelsName.Count();
                    for (int i = LevelsName.Count() + 1; i <= LevelsID.Count(); i++)
                    {
                        var View = ViewPlan.Create(document, ViewFamilyTypeID as ElementId, LevelsID[i] as ElementId);
                        View.Name = $"Level {i}";
                    }
                }
                else if (LevelsName.Count() > LevelsID.Count())
                {
                    for (int i = 0; i < LevelsID.Count(); i++)
                    {
                        var View = ViewPlan.Create(document, ViewFamilyTypeID as ElementId, LevelsID[i] as ElementId);
                        View.Name = LevelsName[i];
                    }
                }
                else
                {
                    MessageBox.Show("There is missing data");
                }
                TnsCreateView.Commit();
            }
        }
        public static void CreateColumn(Document document)
        {
            using (Transaction TnsCreateColumn = new Transaction(document, "Create View"))
            {
                // Get a Column type from Revit
                FilteredElementCollector collector = new FilteredElementCollector(document);
                collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);
                FamilySymbol columnType = collector.FirstElement() as FamilySymbol;

                // Create a level
                Level level = Level.Create(document, 20);

                // Create a FamilyInstance
                FamilyInstance instance;
                XYZ origin = new XYZ(0, 0, 0);
                instance = document.Create.NewFamilyInstance(origin, columnType, level, StructuralType.Column);
                instance.LookupParameter("Width").Set(200);
                instance.LookupParameter("Depth").Set(200);
                instance.LookupParameter("Height").Set(100);
            }
        }
    }
}
