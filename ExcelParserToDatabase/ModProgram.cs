using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbobsCOM;
using System.Threading;

namespace ExcelParserToDatabase
{
    public class ModProgram
    {
        /// <summary>
        /// //////////////////////////////
        /// </summary>

        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.UserObjectsMD oUserObjectMD;
        public static SAPbobsCOM.UserFieldsMD oUserFieldsMD;

        public static int BPLID;


        public static SAPbobsCOM.Recordset oRecordSet;
        public static List<string> ErrMessages;
        private static Users oUser;

        public static string Branches { get; private set; }

        public void Mains(string file, string UDO)
        {

            MainMethod(file, UDO);
        }

        public static void MainMethod(string file, string UDO)
        {
            Application oApp = null;
            oApp = new Application();
            ///////////////////////////////INITIALAZING GLOBAL VARIABLES
            oCompany = TryConnect();
            oUserObjectMD = (SAPbobsCOM.UserObjectsMD)ModProgram.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(ModProgram.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));

            oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery("SELECT MltpBrnchs FROM OADM");
            Branches = oRecordSet.Fields.Item(0).Value.ToString();
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Load migration table addon", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            ExcelParser ep = new ExcelParser();

            ep.FillFromXLFileDI(file, UDO);

            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("load document complited", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }


        public static void Clean()
        {
            //bool success = false;
            for (int i = 0; i < 10; i++)
            {

                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();

                    //Commit transaction
                    if (oCompany.InTransaction)
                    {
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                        GC.Collect();
                        GC.WaitForFullGCComplete();

                        ModProgram.oCompany = TryConnect();



                        break;
                    }


                    Thread.Sleep(100);
                }
            }


        }
        public static Company TryConnect()
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    var comp = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                    return comp;
                }
                catch
                {
                    continue;
                }
            }
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Загрузка не удалась", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            System.Environment.Exit(0);
            return (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();//never going here
        }
        public static void CleanFields()
        {
            ModProgram.oUserFieldsMD = null;
            GC.Collect();
            ModProgram.oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(ModProgram.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));

        }
        public static void CleanUDO()
        {
            ModProgram.oUserObjectMD = null;
            GC.Collect();
            ModProgram.oUserObjectMD = (SAPbobsCOM.UserObjectsMD)ModProgram.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

        }
       
    }
}