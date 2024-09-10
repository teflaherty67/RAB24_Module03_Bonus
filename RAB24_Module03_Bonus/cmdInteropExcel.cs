using Excel = Microsoft.Office.Interop.Excel;

namespace RAB24_Module03_Bonus
{
    [Transaction(TransactionMode.Manual)]
    public class cmdInteropExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // prompt user to select Excel file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            selectFile.InitialDirectory = "S:\\";
            selectFile.Multiselect = false;

            // create an empty string to hold the file name
            string excelFile = "";

            // launch dialog & assign selected file to excelFile if Excel is selected
            if (selectFile.ShowDialog() == DialogResult.OK)
                excelFile = selectFile.FileName;

            // check if Excel file is selected
            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Please select an Excel file.");
                return Result.Failed;
            }

            // if Excel file selected, open it
            Excel.Application excel = new Excel.Application();

            // get the workbook
            Excel.Workbook curWB = excel.Workbooks.Open(excelFile);

            // get the first worksheet
            Excel.Worksheet firstWS = curWB.Worksheets[1];

            // get the range of cells used in worksheet
            Excel.Range range = (Excel.Range)firstWS.UsedRange;

            // get row and column count
            int rows = range.Rows.Count;
            int columnss = range.Columns.Count;

            // read Excel data into a list
            List<List<string>> excelData = new List<List<string>>();
            
            // loop through the rows
            for (int i = 1; i <= rows; i++)
            {
                // create an empty list to hold the row data
                List<string> rowData = new List<string>();
                
                // loop through the columns
                for (int j = 1; j <= columnss; j++)
                {
                    string cellContent = firstWS.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);
                }
                excelData.Add(rowData);
            }

            // save and close Excel
            excel.Save();
            excel.Quit();


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }

}
