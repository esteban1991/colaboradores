using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    class Program
    {

        private static string ItemActiveMenu = null;
        


    
        //instance.SBO_Application_RightClickEvent();
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                   
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }

               
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
                Application.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                
                //Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                //Application.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(sap);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        //private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;

        //    //SAPbouiCOM.Form oForm;
        //    //ItemActiveMenu = pVal.ItemUID;

        //    //try
        //    //{

        //    //    oForm = Application.SBO_Application.Forms.ActiveForm;



        //    //    switch (oForm.TypeEx)
        //    //    {
        //    //        case "Colaboradores_3.Form1":
        //    //            switch (pVal.BeforeAction)
        //    //            {
        //    //                case true:
        //    //                    if (pVal.FormUID == "formpri")
        //    //                    {
        //    //                        Application.SBO_Application.ActivateMenuItem("1289"); //Desactivar Agregar Linea
        //    //                        Application.SBO_Application.ActivateMenuItem("1288"); //Desactivar Borrar Linea
        //    //                    }
        //    //                    else
        //    //                    {
        //    //                        oForm.EnableMenu("1289", true); //Activar Agregar Linea
        //    //                        oForm.EnableMenu("1288", true); //Activar Borrar Linea
        //    //                    }
        //    //                    break;
        //    //            }
        //    //            break;
        //    //    }
        //    //}
        //    //catch (Exception) { }





        //}

   static void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
          SAPbouiCOM.Form oForm;
        
        

         
          ItemActiveMenu= eventInfo.ItemUID;

            
            try
            {

                oForm = Application.SBO_Application.Forms.ActiveForm;



                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.edicionesgr":
                        switch (eventInfo.BeforeAction)
                        {
                            case true:
                                if (eventInfo.ItemUID != "grilaedi")
                                {
                                    oForm.EnableMenu("1292", false); //Desactivar Agregar Linea
                                    oForm.EnableMenu("1293", false); //Desactivar Borrar Linea
                                }
                                else
                                {
                                    oForm.EnableMenu("1292", true); //Activar Agregar Linea
                                    oForm.EnableMenu("1293", true); //Activar Borrar Linea
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception) { }



            try
            {

                oForm = Application.SBO_Application.Forms.ActiveForm;



                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.edimatr":
                        switch (eventInfo.BeforeAction)
                        {
                            case true:
                                if (eventInfo.ItemUID != "Item_2")
                                {
                                    oForm.EnableMenu("1292", false); //Desactivar Agregar Linea
                                    oForm.EnableMenu("1293", false); //Desactivar Borrar Linea
                                }
                                else
                                {
                                    oForm.EnableMenu("1292", true); //Activar Agregar Linea
                                    oForm.EnableMenu("1293", true); //Activar Borrar Linea
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception) { }


            try
            {

                oForm = Application.SBO_Application.Forms.ActiveForm;



                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.suprasec":
                        switch (eventInfo.BeforeAction)
                        {
                            case true:
                                if (eventInfo.ItemUID != "GRSPSEC")
                                {
                                    oForm.EnableMenu("1292", false); //Desactivar Agregar Linea
                                    oForm.EnableMenu("1293", false); //Desactivar Borrar Linea
                                }
                                else
                                {
                                    oForm.EnableMenu("1292", true); //Activar Agregar Linea
                                    oForm.EnableMenu("1293", true); //Activar Borrar Linea
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception) { }



            try
            {

                oForm = Application.SBO_Application.Forms.ActiveForm;



                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.seccionesgr":
                        switch (eventInfo.BeforeAction)
                        {
                            case true:
                                if (eventInfo.ItemUID != "GRSEC")
                                {
                                    oForm.EnableMenu("1292", false); //Desactivar Agregar Linea
                                    oForm.EnableMenu("1293", false); //Desactivar Borrar Linea
                                }
                                else
                                {
                                    oForm.EnableMenu("1292", true); //Activar Agregar Linea
                                    oForm.EnableMenu("1293", true); //Activar Borrar Linea
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception) { }



        }



          static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm;
            SAPbouiCOM.Grid oGrid;
            //(grid) = ItemActiveMenu;
           // string ItemActiveMenu = null;
            SAPbobsCOM.UserTable oUserTable;
             SAPbouiCOM.Application oApp;
            SAPbobsCOM.Company oCompany;
            //SAPbouiCOM.Matrix oMatrix;
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;

                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.edicionesgr":
                        switch (pVal.MenuUID)//1292
                        {
                            case "1292"://Agregar Linea Grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)
                                    {
                                          
                                        case "grilaedi":
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("dted");
                                             DT_GRID.Rows.Add() ; // Agregar una nueva linea
                                             int lastRowIndex = DT_GRID.Rows.Count - 1;
                                             //oGrid.GotFocusAfter.(lastRowIndex);
                                        
                                             //DT_GRID.SetValue("Código", lastRowIndex, "código"); //Asignamos valor a una celda de un nuevo registro
                                             //DT_GRID.SetValue("Nombre", lastRowIndex, "Nombre"); //Asignamos valor a una celda de un nuevo registro
                                             //DT_GRID.SetValue("Proyecto", lastRowIndex, "Proyecto"); //Asignamos valor a una celda de un nuevo registro
                                             //DT_GRID.SetValue("Nombre Proyecto", lastRowIndex, "[Nombre Proyecto]"); //Asignamos valor a una celda de un nuevo registro

                                            break;
                                    }
                                }
                                break;  
                            case "1293"://Eliminar Linea la grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)//obtengo la grilla en el menu actual
                                    {
                                        case "grilaedi":
                                            //aca le indico a mi grid local que sea igual que mi grilla del form activo
                                            oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("grilaedi").Specific);
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("dted");
                                            oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
                                            int nRow = (int)oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                            //obtemos el campo codigo
                                            String sValorGrid = Convert.ToString(oGrid.DataTable.GetValue("Código", nRow));
                                            if (oUserTable.GetByKey(sValorGrid))// verifica si hay un dato 
                                            {
                                                 int mensValor = oApp.MessageBox("¿Esta seguro de eliminar la Edición ?", 1, "Continuar", "Cancelar", "");
                                                //if (System.Windows.Forms.MessageBox.Show("Estas seguro de que quieres eliminar esta sección.", "Advertencia", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                                                 if (mensValor == 1)
                                                 {
                                                     //DT_GRID.Rows.Remove(nRow);
                                                     int i = oUserTable.Remove();


                                                     if (i != 0)
                                                     {
                                                         oApp.SetStatusBarMessage("Error al eliminar : " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                     }
                                                     else
                                                     {
                                                         oApp.SetStatusBarMessage("Edición Eliminada", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                                         //  oForm = ((SAPbouiCOM.Form)oForm.Items.Item("grilaedi").Specific);
                                                         //oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                                         //SAPbouiCOM.Button button2 = ((SAPbouiCOM.Button)oForm.Items.Item("1").Specific);
                                                         //button2.Caption = "out";
                                                     }
                                                 }
                                                 else {
                                                     BubbleEvent = false;
                                                 }
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                        break;
                }
            

            }
            catch (Exception) { }


            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;

                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.suprasec":
                        switch (pVal.MenuUID)//1292
                        {
                            case "1292"://Agregar Linea Grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)
                                    {

                                        case "GRSPSEC":
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DTSPSEC");
                                            DT_GRID.Rows.Add(); // Agregar una nueva linea
                                            int lastRowIndex = DT_GRID.Rows.Count - 1;
                                            //oGrid.GotFocusAfter.(lastRowIndex);

                                            //DT_GRID.SetValue("Código", lastRowIndex, "código"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Nombre", lastRowIndex, "Nombre"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Proyecto", lastRowIndex, "Proyecto"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Nombre Proyecto", lastRowIndex, "[Nombre Proyecto]"); //Asignamos valor a una celda de un nuevo registro

                                            break;
                                    }
                                }
                                break;
                            case "1293"://Eliminar Linea la grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)//obtengo la grilla en el menu actual
                                    {
                                        case "GRSPSEC":
                                            //aca le indico a mi grid local que sea igual que mi grilla del form activo
                                            oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("GRSPSEC").Specific);
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DTSPSEC");
                                            oUserTable = oCompany.UserTables.Item("SUPRASECCIONESCOL");
                                            int nRow = (int)oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                            //obtemos el campo codigo
                                            String sValorGrid = Convert.ToString(oGrid.DataTable.GetValue("Código", nRow));
                                            if (oUserTable.GetByKey(sValorGrid))// verifica si hay un dato 
                                            {
                                                  int mensValor = oApp.MessageBox("¿Esta seguro de eliminar la Supra-sección ?", 1, "Continuar", "Cancelar", "");
                                                //if (System.Windows.Forms.MessageBox.Show("Estas seguro de que quieres eliminar esta sección.", "Advertencia", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                                                  if (mensValor == 1)
                                                  {

                                                      //DT_GRID.Rows.Remove(nRow);
                                                      int i = oUserTable.Remove();


                                                      if (i != 0)
                                                      {
                                                          oApp.SetStatusBarMessage("Error al eliminar : " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                      }
                                                      else
                                                      {
                                                          oApp.SetStatusBarMessage("Supra-Sección Eliminada", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                                          //  oForm = ((SAPbouiCOM.Form)oForm.Items.Item("grilaedi").Specific);
                                                          //oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                                          //SAPbouiCOM.Button button2 = ((SAPbouiCOM.Button)oForm.Items.Item("1").Specific);
                                                          //button2.Caption = "out";
                                                      }
                                                  }
                                                  else {
                                                      BubbleEvent = false;
                                                  }
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                        break;
                }


            }
            catch (Exception) { }





            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;

                switch (oForm.TypeEx)
                {
                    case "Colaboradores_3.seccionesgr":
                        switch (pVal.MenuUID)//1292
                        {
                            case "1292"://Agregar Linea Grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)
                                    {

                                        case "GRSEC":
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DTSEC");
                                            DT_GRID.Rows.Add(); // Agregar una nueva linea
                                            int lastRowIndex = DT_GRID.Rows.Count - 1;
                                            //oGrid.GotFocusAfter.(lastRowIndex);

                                            //DT_GRID.SetValue("Código", lastRowIndex, "código"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Nombre", lastRowIndex, "Nombre"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Proyecto", lastRowIndex, "Proyecto"); //Asignamos valor a una celda de un nuevo registro
                                            //DT_GRID.SetValue("Nombre Proyecto", lastRowIndex, "[Nombre Proyecto]"); //Asignamos valor a una celda de un nuevo registro

                                            break;
                                    }
                                }
                                break;
                            case "1293"://Eliminar Linea la grilla
                                if (pVal.BeforeAction == true)
                                {
                                    switch (ItemActiveMenu)//obtengo la grilla en el menu actual
                                    {
                                        case "GRSEC":
                                            //aca le indico a mi grid local que sea igual que mi grilla del form activo
                                            oGrid = ((SAPbouiCOM.Grid)oForm.Items.Item("GRSEC").Specific);
                                            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DTSEC");
                                            oUserTable = oCompany.UserTables.Item("SECCIONESCOL");
                                            int nRow = (int)oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                            //obtemos el campo codigo
                                            String sValorGrid = Convert.ToString(oGrid.DataTable.GetValue("Código", nRow));
                                            if (oUserTable.GetByKey(sValorGrid))// verifica si hay un dato 
                                            {
                                                //le pregunto si realmente desea eliminar la fila
                                               int mensValor = oApp.MessageBox("¿Esta seguro de eliminar la sección ?", 1, "Continuar", "Cancelar", "");
                                                //if (System.Windows.Forms.MessageBox.Show("Estas seguro de que quieres eliminar esta sección.", "Advertencia", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                                               if (mensValor == 1)
                                                {
                                                    //DT_GRID.Rows.Remove(nRow);
                                                    int i = oUserTable.Remove();


                                                    if (i != 0)
                                                    {
                                                        oApp.SetStatusBarMessage("Error al eliminar : " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                    }
                                                    else
                                                    {
                                                        oApp.SetStatusBarMessage("Sección Eliminada", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                                                    }
                                                }
                                                else {

                                                    BubbleEvent = false;
                                                }
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                        break;
                }


            }
            catch (Exception) { }





            //try
            //{
            //    oForm = Application.SBO_Application.Forms.ActiveForm;

            //    switch (oForm.TypeEx)
            //    {
            //        case "Colaboradores_3.edimatr":
            //            switch (pVal.MenuUID)//1292
            //            {
            //                case "1292"://Agregar Linea Matrix
            //                    if (pVal.BeforeAction == true)
            //                    {
            //                        switch (ItemActiveMenu)
            //                        {

            //                            case "Item_2":

            //                                AddLineMatrixDBDataSource(((SAPbouiCOM.Matrix)oForm.Items.Item("Item_2").Specific), oForm.DataSources.DBDataSources.Item("@EDICIONESCOL"), "U_CodigoEDC");
            //                                break;
            //                                //SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("dted");
            //                                //DT_GRID.Rows.Add(); // Agregar una nueva linea
            //                                //int lastRowIndex = DT_GRID.Rows.Count - 1;

            //                                //DT_GRID.SetValue("Código", lastRowIndex, "código"); //Asignamos valor a una celda de un nuevo registro
            //                                //DT_GRID.SetValue("Nombre", lastRowIndex, "Nombre"); //Asignamos valor a una celda de un nuevo registro
            //                                //DT_GRID.SetValue("Proyecto", lastRowIndex, "Proyecto"); //Asignamos valor a una celda de un nuevo registro
            //                                //DT_GRID.SetValue("Nombre Proyecto", lastRowIndex, "[Nombre Proyecto]"); //Asignamos valor a una celda de un nuevo registro

                                           
            //                        }
            //                    }
            //                    break;
            //                case "1293"://Eliminar Linea Matrix
            //                    if (pVal.BeforeAction == true)
            //                    {
            //                        switch (ItemActiveMenu)
            //                        {
            //                            case "Item_2":

            //                        //        SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("grilaedi");
            //                        //        //   DT_GRID.Rows.Remove(NumeroRow);
            //                        //        break;
                                            
            //                                oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("Item_2").Specific);
            //                                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@EDICIONESCOL");
            //                                int nRow = (int)oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            //                                oMatrix.FlushToDataSource();
            //                                oDBDataSource.RemoveRecord(nRow-1);
            //                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            //                                oMatrix.LoadFromDataSource();
            //                                BubbleEvent = false;
            //                                break;
            //                        }
            //                    }
            //                    break;
            //            }
            //            break;
            //    }
            //}
            //catch (Exception) { }

          
        }

          public static void AddLineMatrixDBDataSource(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource source, string ColumnaFocus = "")
          {
              try
              {
                  SAPbouiCOM.EditText oEdit;
                  oMatrix.FlushToDataSource();
                  source.InsertRecord(source.Size);
                  source.Offset = source.Size - 1;
                  oMatrix.LoadFromDataSource();
                  oMatrix.SelectRow(source.Size, true, false);
                  if (ColumnaFocus.Trim().Length != 0)
                  {
                      oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(ColumnaFocus).Cells.Item(source.Size).Specific;
                      oEdit.Active = true;
                      oEdit.Item.Enabled = true;
                  }
              }
              catch (Exception) { }
          }


         public class tablas
          {
              public static SAPbouiCOM.Application oApp;
              public static SAPbobsCOM.Company oCompany;

              public static string CreateUDT(string tableName, string tableDesc, SAPbobsCOM.BoUTBTableType tableType)
              {
                  oApp = (SAPbouiCOM.Application)Application.SBO_Application;
                  oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
                  SAPbobsCOM.UserTablesMD oUdtMD = null;
                  SAPbobsCOM.UserFieldsMD oUdtCA = null;
                  string VERSION= null;
                   SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                   oRec.DoQuery(" SELECT [U_APLICACION] as 'SAP',[U_Version]  As 'Version',[U_fecha] as Fecha FROM [@VERSION]");

                        oRec.MoveFirst();
                        while (!oRec.EoF)
                        {
                            //oCBC.ValidValues.Add(oRec.Fields.Item(0).Value.ToString(), oRec.Fields.Item(1).Value.ToString());
                            VERSION = oRec.Fields.Item("Version").Value.ToString();

                           // Extraerdequery1 = Convert.ToString(Grid2.Columns.Item("Descripción"));
                
                            oRec.MoveNext();
               
                        }
                        if (VERSION == "1.1")
                        {

                            tableName = "@Prueba";
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


                                    return "exito";
                                }
                                else if (VERSION == "1.2")
                                {



                                    return "";
                                }

                                else {

                                    return "";
                                }
                            }
                            catch (Exception ex)
                            {
                                return ex.Message;
                            }
                        }
                        else {

                            return "";
                        }
              }
              
          }



        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

    
    }
}
