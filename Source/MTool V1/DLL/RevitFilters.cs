namespace MTool_V1.DLL
{
    using System;
    using System.Activities.Expressions;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Workflow.ComponentModel;
    using Autodesk.AutoCAD.BoundaryRepresentation;
    using Autodesk.AutoCAD.Customization;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Element = Autodesk.Revit.DB.Element;

    public static class RevitFilters
    {
        private const string filepath = @".\Excel Files\Creating.xlsx";
        private const string logElementNameWithIdFilepath = @".\Logs\LogElementNameWithIDFile.txt";
        private const string logGetElementIdFilepath = @".\Logs\logGetElementIdFile.txt";
        private const string logGetElementNameFilepath = @".\Logs\logGetElementNameFile.txt";
        private const string logViewFamilyTypeNameWithIdFilepath = @".\Logs\LogViewFamilyTypeNameWithIDFile.txt";
        private const string logGetElementsDataByClassFilepath = @".\Logs\LogGetElementsDataByClassFile.txt";
        private const string logGetElementsDataByClassAndCategoryFilepath = @".\Logs\LogGetElementsDataByClassAndCategoryFile.txt"; 
        private const string logGetFamilyInstanceByTypeOfFamilyFilepath = @".\Logs\LogGetFamilyInstanceByTypeOfFamilyFile.txt";
        
        public enum ReturnElementData
        {
            ElementID,
            ElementName
        }
        public enum ReturnViewData
        {
            ViewFamilyTpeID,
            ViewFamilyTpeName
        }
        /// <summary>
        /// Get All Elements Without adding any Filter Data as list of (ID <ElementID> Or Name <string>). 
        /// </summary>
        /// <param name="document">The decoumant that include the database.</param>
        /// <param name="builtInCategory">The Catecory Type want to get the data from it.</param>
        /// <param name="returnData">ElemantID as list<ElementID> Or ElemnetName as list<string>.</param>
        /// <param name="ElementReturnName">This for Log File Only.</param>
        /// <returns></returns>
        public static object GetAllElementsDataByCategory(Document document, BuiltInCategory builtInCategory, ReturnElementData returnData, string ElementReturnName = "Levels")
        {
            var Logfile = new StringBuilder();
            Logfile.Append($"Successed To Get {ElementReturnName}:\n");
            var CollectorElementID = new FilteredElementCollector(document);
            var ELements = CollectorElementID.OfCategory(builtInCategory)
                .ToElements();
            var ELementsID = new List<ElementId>(); 
            var ELementsName =new List<string>();
            int index = 0;
            foreach (var Element in ELements)
            {
                ELementsID.Add(Element.Id);
                ELementsName.Add(Element.Name);
                Logfile.Append($"Name: {ELementsName[index]} ----- ID: {ELementsID[index].ToString()}\n");
                index++;
            }
            Open_StreamFiles.CreatEncryptedLogFile(logElementNameWithIdFilepath, ".log", Logfile, "Revit Filters", false);
            switch (returnData)
            {
                case ReturnElementData.ElementID:
                    return ELementsID;
                default:
                    return ELementsName;
            }
        }
        /// <summary>
        /// Get Elements Data from Given list and return Data as list of (ID <ElementID> Or Name <string>). 
        /// </summary>
        /// <param name="document">The decoumant that include the database.</param>
        /// <param name="builtInCategory">The Catecory Type want to get the data from it.</param>
        /// <param name="list">Given List<string>.</param>
        /// <param name="ElementReturnData">This for Log File Only.</param>
        /// <returns></returns>
        public static object GetElementsDataByList(Document document,BuiltInCategory builtInCategory, List<string> list, string ElementReturnData = "Levels ID")
        {
            var Logfile = new StringBuilder();
            Logfile.Append($"Successeded To Get {ElementReturnData}:\n");
            var ELementsID = new List<ElementId>();
            var Elements = new List<Autodesk.Revit.DB.Element>();
            var ELementsName = new List<string>();

            foreach (var NameOfElement in list)
            {
                var collector = new FilteredElementCollector(document);
                Elements = collector
                    .OfCategory(builtInCategory).ToElements()
                    .Where(x => x.Name == NameOfElement).ToList();
                foreach (var item in Elements)
                {
                    ELementsID.Add(item.Id);
                    Logfile.Append($"ID: {item.Id.ToString()} ----- Name: {NameOfElement}\n");
                }
            }
            Open_StreamFiles.CreatEncryptedLogFile(logGetElementIdFilepath, ".log", Logfile, "Revit Filters", false);
            return ELementsID;
        }
        /// <summary>
        /// Get Elements Data from Given list and return Data as list of (ID <ElementID> Or Name <string>). 
        /// </summary>
        /// <param name="document">The decoumant that include the database.</param>
        /// <param name="builtInCategory">The Catecory Type want to get the data from it.</param>
        /// <param name="list">Given List<ElementId>.</param>
        /// <param name="ElementReturnData">This for Log File Only.</param>
        /// <returns></returns>
        public static object GetElementsDataByList(Document document, BuiltInCategory builtInCategory, List<ElementId> list, string ElementReturnData = "Levels Name")
        {
            var Logfile = new StringBuilder();
            Logfile.Append($"Successed To Get {ElementReturnData}:\n");
            var ELementsID = new List<ElementId>();
            var Elements = new List<Autodesk.Revit.DB.Element>();
            var ELementsName = new List<string>();

            foreach (var Element in list)
            {
                var collector = new FilteredElementCollector(document);
                Elements = collector.OfCategory(builtInCategory).ToElements()
                            .Where(x => x.Id == Element).ToList();
                foreach (var item in Elements)
                {
                    ELementsName.Add(item.Name);
                    Logfile.Append($"Name: {item.Name} ----- ID: {Element.ToString()}\n");
                }
            }
            Open_StreamFiles.CreatEncryptedLogFile(logGetElementNameFilepath, ".log", Logfile, "Revit Filters", false);
            return ELementsName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="returnData"></param>
        /// <returns></returns>
        public static object GetViewFamilyType(Document document , ReturnViewData returnData)
        {
            var Logfile = new StringBuilder();
            var CollectorViewFamilyTypeID = new FilteredElementCollector(document);
            var ViewFamilyType = CollectorViewFamilyTypeID.OfClass(typeof(ViewFamilyType))
                .FirstOrDefault();
            var ViewFamilyTypeName = ViewFamilyType.Name;
            var ViewFamilyTypeID = ViewFamilyType.Id;
            Logfile.Append($"Successed Get ViewFamilyType:\nName: {ViewFamilyTypeName} ----- ID: {ViewFamilyTypeID.ToString()}\n");
            Open_StreamFiles.CreatEncryptedLogFile(logViewFamilyTypeNameWithIdFilepath, ".log", Logfile, "Revit Filters", false);
            switch (returnData)
            {
                case ReturnViewData.ViewFamilyTpeID:
                    return ViewFamilyTypeID;
                default:
                    return ViewFamilyTypeName;
            }
        }
        public static object GetElementsDataByClass(Document document, Type type, ReturnElementData returnData,BuiltInCategory builtInCategory , string ElementReturnName = "Families Instance")
        {
            var Logfile = new StringBuilder();
            Logfile.Append($"Successed To Get {ElementReturnName}:\n");
            var CollectorElementID = new FilteredElementCollector(document);
            var ELements = CollectorElementID.OfClass(type)
                .OfCategory(builtInCategory)
                .ToElements();
            var ELementsID = new List<ElementId>();
            var ELementsName = new List<string>();
            int index = 0;
            foreach (var Element in ELements)
            {
                ELementsID.Add(Element.Id);
                ELementsName.Add(Element.Name);
                Logfile.Append($"Name: {ELementsName[index]} ----- ID: {ELementsID[index].ToString()}\n");
                index++;
            }
            Open_StreamFiles.CreatEncryptedLogFile(logGetElementsDataByClassFilepath, ".log", Logfile, "Revit Filters", false);
            switch (returnData)
            {
                case ReturnElementData.ElementID:
                    return ELementsID;
                default:
                    return ELementsName;
            }
        }
        public static object GetElementsDataByClass(Document document, Type type, ReturnElementData returnData, string ElementReturnName = "Families")
        {
            var Logfile = new StringBuilder();
            Logfile.Append($"Successed To Get {ElementReturnName}:\n");
            var CollectorElementID = new FilteredElementCollector(document);
            var ELements = CollectorElementID.OfClass(type)
                .ToElements();
            var ELementsID = new List<ElementId>();
            var ELementsName = new List<string>();
            int index = 0;
            foreach (var Element in ELements)
            {
                ELementsID.Add(Element.Id);
                ELementsName.Add(Element.Name);
                Logfile.Append($"Name: {ELementsName[index]} ----- ID: {ELementsID[index].ToString()}\n");
                index++;
            }
            Open_StreamFiles.CreatEncryptedLogFile(logGetElementsDataByClassAndCategoryFilepath, ".log", Logfile, "Revit Filters", false);
            switch (returnData)
            {
                case ReturnElementData.ElementID:
                    return ELementsID;
                default:
                    return ELementsName;
            }
        }
        public static object GetFamilyInstanceByTypeOfFamily(Document document, string FamilyName, ReturnElementData returnData, string ElementReturnName = "Family Instance")
        {
            var Logfile = new StringBuilder();
            var Familylist = new List<Element>();
            var EnumElement = RevitConveter.GetEnumInBuiltInCategoryByText(FamilyName);
                var CollectorCategory = new FilteredElementCollector(document);
            Familylist = CollectorCategory
                .OfCategory(BuiltInCategory.OST_StructuralColumns).ToElements().ToList();
            var ELementsID = new List<ElementId>();
            var ELementsName = new List<string>();
            int index = 0;
            foreach (var Element in Familylist)
            {
                ELementsID.Add(Element.Id);
                ELementsName.Add(Element.Name);
                Logfile.Append($"{index+1}- Name: {ELementsName[index]} ----- ID: {ELementsID[index].ToString()}\n");
                index++;
            }
            Open_StreamFiles.CreatEncryptedLogFile(logGetFamilyInstanceByTypeOfFamilyFilepath, ".log", Logfile, "Revit Filters", false);
            switch (returnData)
            {
                case ReturnElementData.ElementID:
                    return ELementsID;
                default:
                    return ELementsName;
            }
        }
    }
}
