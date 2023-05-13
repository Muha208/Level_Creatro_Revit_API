using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Windows.Input;

namespace MTool_V1.Commands
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class MtoolsCommands : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                //to make it an active veiw
                Main_Window_MTools windowM = new Main_Window_MTools(doc);
                windowM.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
                return Result.Failed;
            }
        }
    }
}