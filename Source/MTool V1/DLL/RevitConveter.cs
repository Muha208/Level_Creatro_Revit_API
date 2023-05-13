namespace MTool_V1.DLL
{
    using Autodesk.Revit.DB;
    using Spire.Pdf.Exporting.XPS.Schema;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public static class RevitConveter
    {
        public static Enum GetEnumInBuiltInCategoryByText(string EnumName)
        {
            var EnumArray = Enum.GetValues(typeof(BuiltInCategory));
            var EnumList = EnumArray.Cast<Enum>().ToList();
            var NameOfCatecory = $"{EnumName}";
            for (int i = 0; i < EnumList.Count; i++)
            {
                if (EnumList[i].ToString() == NameOfCatecory)
                {
                    return EnumList[i];
                }
            }
            return null;
        }
    }
}
