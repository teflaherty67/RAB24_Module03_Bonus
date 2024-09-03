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
