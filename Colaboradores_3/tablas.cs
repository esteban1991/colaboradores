using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
namespace Colaboradores_3
{
    class tablas
    {
        public static SAPbouiCOM.Application oApp;
        public static SAPbobsCOM.Company oCompany;

        public static string CreateUDT(string tableName, string tableDesc, SAPbobsCOM.BoUTBTableType tableType)
        {
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            SAPbobsCOM.UserTablesMD oUdtMD = null;
            SAPbobsCOM.UserFieldsMD oUdtCA = null;
            tableName = "Prueba";
            try
            {
                oUdtMD = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUdtMD.GetByKey(tableName) == false)
                {

                    oUdtMD.TableName = tableName;
                    oUdtMD.TableDescription = tableName;
                    oUdtMD.TableType = tableType;
                    
                   
             
                    
                   // oUdtCA.Mandatory = "tYes";
                    
                    int lRetCode;
                    lRetCode = oUdtMD.Add();
                    oUdtCA.Add();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdtMD);
                    oUdtMD = null;
                    GC.Collect();
                    if ((lRetCode != 0))
                    {



                        if ((lRetCode == -2035))
                        {
                            return "-2035";
                        }

                        return oCompany.GetLastErrorDescription();
                    }


                    return "";
                }
                else
                {

              
                                       
                    return "";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

    }
}
