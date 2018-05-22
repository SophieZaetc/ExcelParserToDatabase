using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParserToDatabase
{
    class ExcelParser
    {
        public void FillFromXLFileDI(string pathWithName,string UDOName)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Microsoft.Office.Interop.Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pathWithName); //pathWithName here
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            var ColumnList = new List<string>();
            for (int j = 1; j <= colCount; j++)
            {
                ColumnList.Add(xlRange.Cells[1, j].Value2.ToString());
            }

            for (int i = 2; i <= rowCount; i++)
            {
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralData;
                var oCompService = ModProgram.oCompany.GetCompanyService();
                ModProgram.oCompany.StartTransaction();
                oGeneralService = oCompService.GetGeneralService(UDOName);
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                for (int j = 0; j <= ColumnList.Count; j++)
                {

                    // i rowcount 
                    // j collcount
                    //write the value to the console
                    if (xlRange.Cells[i, j + 1] != null && xlRange.Cells[i, j + 1].Value2 != null)
                    {

                        try
                        {
                            var date = (DateTime.FromOADate(xlRange.Cells[i, j + 1].Value2));
                            oGeneralData.SetProperty(ColumnList[j], date);
                        }

                        catch
                        {
                            oGeneralData.SetProperty(ColumnList[j], xlRange.Cells[i, j + 1].Value2.ToString().Replace("'", "").Replace(",", "."));
                        } 
                    string buuf = xlRange.Cells[i, j + 1].Value2.ToString().Replace("'", "").Replace(",", ".");
                    string buffer = xlRange.Cells[i, j + 1].Value2.ToString().Replace("'", "").Replace(",", "."); Debug.Write(buffer);
                        ModProgram.oRecordSet.DoQuery("select isnull(max(DocEntry)+1,1) from [@BDO_UKR_OMTD]");
                        oGeneralData.SetProperty("Code", ModProgram.oRecordSet.Fields.Item(0).Value.ToString());
                    }

            }
            oGeneralService.Add(oGeneralData);
            if (ModProgram.oCompany.InTransaction)
            {
                ModProgram.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
        }

        //cleanup
        GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

}
}
