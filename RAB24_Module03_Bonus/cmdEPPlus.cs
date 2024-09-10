using Microsoft.Office.Interop.Excel;

namespace RAB24_Module03_Bonus
{
    [Transaction(TransactionMode.Manual)]
    public class cmdEPPlus : IExternalCommand
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

            // set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // if Excel file selected, open it
            ExcelPackage excel = new ExcelPackage(excelFile);

            // get the workbook
            ExcelWorkbook curWB = excel.Workbook;

            // get the first worksheet
            ExcelWorksheet firstWS = curWB.Worksheets[0];

            // get row and column count
            int rows = firstWS.Dimension.Rows;
            int columnss = firstWS.Dimension.Columns;

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
            excel.Dispose();


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData.Data;
        }
    }

}
