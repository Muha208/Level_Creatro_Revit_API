using System;
using System.Reflection;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Security.Policy;
using System.IO.Packaging;
using System.ServiceModel.Channels;
using System.Windows;
using System.IO;

namespace MTool_V1
{
    public class Main : IExternalApplication
    {
        
        #region external application public methods
   
        public Result OnStartup(UIControlledApplication application)
        {

            string FilepathImgMAin = @"M:\Work\Coding\KAITECH Assigment\Assigment\KAITECH Assignments\KAITECH Assignments\Assignments\Level Creator\Source\MTool V1\bin\Debug\MIcons\MTools Main.png";
            string FilepathImgLarg = @"M:\Work\Coding\KAITECH Assigment\Assigment\KAITECH Assignments\KAITECH Assignments\Assignments\Level Creator\Source\MTool V1\bin\Debug\MIcons\MTools Large.png";

            //plugin's main Tab name (the name of tab that will inculd your tools)
            string tabName = "MTools";
            //panel name hosted on ribbon tab (descreption of all tools that inside your rebbon)
            string panelAnnotationName = ("KAUTECH");
            //craet tab on Revit UI
            application.CreateRibbonTab(tabName);
            //Creat panel on Revit rebbon tab (that the panel opend after click on rebbon and it will apear the tools)
            var panelAnnotation = application.CreateRibbonPanel(tabName,panelAnnotationName);
            //Create push buttom and populate it with information
            //Location = the namespace.Classname
            //it's need (the tab that you want to put it in it, the name of your button that you need to apear, ,The class file (namespace+main class) that include your code of this pushbutton)
            //this code for insert data only
            var mTools = new PushButtonData(tabName, "Create Levels", Assembly.GetExecutingAssembly().Location, "MTool_V1.Commands.MtoolsCommands")
            {
                //This is the Bitmap Image will appeared in Rebbon (small one)
                 //ToolTipImage = new BitmapImage(new Uri(FilepathImgMAin)),
                 ToolTip = "KAITECH"
            };
            //but this this code to create the pushbutton that will include your data
            //this is the main bitmap (larg 350x350 px)
            var mToolslayer = panelAnnotation.AddItem(mTools) as PushButton;
            mToolslayer.ItemText = "Create Levels"; 
            //mToolslayer.LargeImage = new BitmapImage(new Uri(FilepathImgLarg));

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        #endregion
    }
}
