using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.IO;
using System.Web.UI.WebControls;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using MTool_V1.DLL;
using MessageBox = System.Windows.MessageBox;
using System.Reflection;

namespace MTool_V1
{

    /// <summary>
    /// Interaction logic for Main_Window_MTools.xaml
    /// </summary>
    /// 

    public partial class Main_Window_MTools : System.Windows.Window 
     {
        Document document;
        public Main_Window_MTools(Document doc)
        {
            InitializeComponent();
            document = doc;    
        }


        //varab
        public OpenFileDialog ofd = null;
        public int ofdTrue = 0;
        public FileStream stream = null;
        public string filePath;
        public FileSaveDialog fileSaveRevit = null;

        public string pathfileget(string path, int ofdTrue)
        {
            if (path == null && ofdTrue == 0)
            {
                path = @"M:\Work\Coding\C# Projects\MTools\Excell Files\Creatingf.xls";
                return path;
            }
            else
                return path;
        }
        public void openFileDialogEX()
        {
            ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx";
            ofd.Multiselect = false;
        }

        public void Close_B_Tip(object sender, MouseButtonEventArgs e)
        {
            this.Close();
            stream.Close();
        }
        public void Mini_Eve(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        public void Git_About(object sender, MouseButtonEventArgs e)
        {
            //for Logo Image to open the URL of my repo on github site to more info
            this.WindowState = WindowState.Minimized;
            var giT = Uri.UriSchemeHttps;
            giT = "https://github.com/Muha208/MTools-Revit-API";
            Process.Start(giT);
        }

        public void Click_Creat_But(object sender, RoutedEventArgs e)
        {
            MainGridSC.Visibility = System.Windows.Visibility.Hidden;
            MainGridCF.Visibility = System.Windows.Visibility.Visible;
        }
        public void MaxBot(object sender, MouseButtonEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }else
            this.WindowState = WindowState.Maximized;
        }
        private void ImpFile_but(object sender, RoutedEventArgs e)
        {
            //Create Levels with Floor Plan Views (From Excel File)
            try
            {
                var LevelsName = new List<string>();
                LevelsName = Open_StreamFiles.GetColumnTextDataFromExcelFile("Levels").ToList();
                var LevelsNumber = new List<double>();
                LevelsNumber = Open_StreamFiles.GetColumnNumericalDataFromExcelFile("Levels", FixedColumns: 3).ToList();
                OwneRevitMathods.CreateLevels(document, LevelsNumber, LevelsName);
                var LevelsID = new List<ElementId>();
                LevelsID = (List<ElementId>)RevitFilters.GetElementsDataByList(document, BuiltInCategory.OST_Levels, LevelsName);
                OwneRevitMathods.CreateViews(document, LevelsID, LevelsName);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int indexsheet = SheetsCombo.SelectedIndex;
            ExcelSheets.ActiveSheetIndex = indexsheet;
        }

        private void Tab1Loaded(object sender, RoutedEventArgs e)
        {
            #region Streaming
            try
            {
                using (stream = Open_StreamFiles.OpenStreamExcelFile())
                {
                    if (ExcelSheets.Worksheets.Count != 0)
                    {
                        for (int i = 0; i < ExcelSheets.Worksheets.Count; i++)
                        {
                            if (ExcelSheets.Worksheets[i].SheetName == "Evaluation Warning")
                            {
                                ExcelSheets.Worksheets.Remove(i);
                            }
                        }
                    }
                    ExcelSheets.LoadFromFile(stream.Name, true);
                    int ExsheetCount = ExcelSheets.Worksheets.Count;
                    if (ExsheetCount > 0)
                    {
                        SheetsCombo.Items.Clear();
                        for (int i = 0; i < ExsheetCount; i++)
                        {
                            string addsheet = ExcelSheets.Worksheets[i].SheetName;
                            SheetsCombo.Items.Add(addsheet);
                        }

                    }
                    ExcelSheets.ActiveSheetIndex = 0;
                    SheetsCombo.SelectedIndex = 0;
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
            #endregion
            
        }


        private void mousedown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void openEX_Click(object sender, RoutedEventArgs e)
        {
            openFileDialogEX();
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = ofd.FileName;
                ofdTrue = 1;
                if (stream != null)
                {
                    stream.Close();
                    stream = new FileStream(pathfileget(filePath, ofdTrue), FileMode.Open, FileAccess.ReadWrite);
                }
                else
                {
                    stream = new FileStream(pathfileget(filePath, ofdTrue), FileMode.Open, FileAccess.ReadWrite);
                }
                ExcelSheets.ActiveSheetIndex = 0;
                int ExsheetCount = ExcelSheets.Worksheets.Count;
                SheetsCombo.Items.Clear();
                for (int i = 0; i < ExsheetCount; i++)
                {
                    string addsheet = ExcelSheets.Worksheets[i].SheetName;
                    SheetsCombo.Items.Add(addsheet);
                }
                SheetsCombo.SelectedIndex = 0;
            }
            else
                System.Windows.MessageBox.Show("Please Select Execl File");

        }
    }

}
