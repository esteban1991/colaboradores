using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {

        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Grid oGrid;

        public Form1()
        {


        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("TxtCod").Specific));
            this.EditText0.ClickAfter += new SAPbouiCOM._IEditTextEvents_ClickAfterEventHandler(this.EditText0_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("TxtNom").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("TxtNIF").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lblnif").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("TxtAli").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("pesdgenrl").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("pesdeco").Specific));
            this.Folder1.ClickBefore += new SAPbouiCOM._IFolderEvents_ClickBeforeEventHandler(this.Folder1_ClickBefore);
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Txt_direc").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("TxtPobl").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_17").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("TxtCP").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("TxtPro").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("TxtTel1").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_25").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_26").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("TxtTel2").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("TxtFax").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_30").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("TxtEmail").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbPais").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbfpago").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_5").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmpgbc").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_24").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbctbc").Specific));
            this.ComboBox3.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox3_ComboSelectBefore);
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("txtsucur").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_32").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_33").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_35").Specific));
            this.ComboBox4 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbcp").Specific));
            //                                         this.ComboBox4.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox4_ComboSelectAfter);
            this.ComboBox4.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox4_ComboSelectBefore);
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("pesobs").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("bgtxtfree").Specific));
            this.Folder6 = ((SAPbouiCOM.Folder)(this.GetItem("Item_44").Specific));
            this.oGrid = ((SAPbouiCOM.Grid)(this.GetItem("grdpy").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("TXTIBAN").Specific));
            this.ComboBox5 = ((SAPbouiCOM.ComboBox)(this.GetItem("CMBIBAN").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("327").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("txtbics").Specific));
            this.StaticText20 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("txtcotbn").Specific));
            this.OnCustomInitialize();

        }

  

       

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        /// 




        //CLASE PARA LIMPIAR LOS COMBOBOXS
        public static void CleanComboBox(dynamic oComboBox)
        {
            int i = 0;

            while (oComboBox.ValidValues.Count > 0)
            {
                try
                {
                    oComboBox.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }

            }
        }

        //cargar  por recorsert los combobox  utilizando los validvalues
        private void LoadComboQueryRecordset(string _query, dynamic oComboBox, string fieldValue, string fieldDesc, SAPbobsCOM.Company oCompany)
        {

            //oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            //oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();

            SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            businessObject.DoQuery(_query);
            SAPbouiCOM.ValidValues validValues = oComboBox.ValidValues;

            CleanComboBox(oComboBox);

            if (!string.Equals(fieldDesc, string.Empty))
            {
                while (!businessObject.EoF)
                {
                    validValues.Add((dynamic)businessObject.Fields.Item(fieldValue).Value, (dynamic)businessObject.Fields.Item(fieldDesc).Value);
                    businessObject.MoveNext();
                }
            }
            else
            {
                while (!businessObject.EoF)
                {
                    validValues.Add((dynamic)businessObject.Fields.Item(fieldValue).Value, "");
                    businessObject.MoveNext();
                }
            }
        }

        public override void OnInitializeFormEvents()
        {

        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {
            //conexion
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //combobox de paises 
            CleanComboBox(ComboBox2);
            string SqlCad = "select code,name from ocry";
            LoadComboQueryRecordset(SqlCad, ComboBox2, "code", "name", oCompany);
            //Folder0.Item.Click();
            //combobox de f_pagos 
            CleanComboBox(ComboBox0);
            string SqlCad2 = "select GroupNum,PymntGroup from octg";
            
            LoadComboQueryRecordset(SqlCad2, ComboBox0, "GroupNum", "PymntGroup", oCompany);

            //combobox de pais de los bancos
            CleanComboBox(ComboBox4);
            string SqlCad5 = "select DISTINCT ocry.code,ocry.name from ocry  inner join odsc on ocry.code=odsc.CountryCod ";
            LoadComboQueryRecordset(SqlCad5, ComboBox4, "code", "name", oCompany);
             


            oForm = oApp.Forms.Item("formpri");
            //  Muestra el formulario
            oForm.Visible = true;
            //para que funcionen las flechas y e refresh (databrowser)
            oForm.DataBrowser.BrowseBy = "TxtCod";
            //busca la grilla con el UID
         
            oitem = oForm.Items.Item("grdpy");
            oGrid = ((SAPbouiCOM.Grid)(oitem.Specific));
            

            //crear el datasources  y despues dentro de esto se ejecuta la query
            oForm.DataSources.DataTables.Add("grPagdt");

            oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT distinct t0.PayMethCod AS Código,t0.Descript as Descripción,t0.Active as Activo from OPYM  as t0 left join  ocrd as t1  on t0.PayMethCod= t1.pymcode where t0.type='O' ");
                oGrid.DataTable = oForm.DataSources.DataTables.Item("grPagdt");

            //cheackbox en el campo activo 
            oGrid.Columns.Item("Activo").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            //para agregar el link buton a la grilla
            oGrid.Columns.Item("Código").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
            SAPbouiCOM.EditTextColumn oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Código");
            oEdit.LinkedObjectType = "147";
            // oGrid.Rows.SelectedRows.Add(pVal.Row);
            //string sMetodoPago = (string)oGrid.Columns.Item("Activo");
            //int nRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            //string sMetodoPago = (string)oGrid.DataTable.GetValue("Código", nRow);
      

            RowNumberGrid(oGrid);
         
        }

        //metodo: el cual recibe como parametro el objeto grid que quieras numerar
        public static void RowNumberGrid(SAPbouiCOM.Grid oGrid)
        {
            SAPbouiCOM.RowHeaders oHeader = null;
            oHeader = oGrid.RowHeaders;

            for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
            {
                oHeader.SetText(i, Convert.ToString(i + 1));

            }
        }
        public void grillametpag() {

            oitem = oForm.Items.Item("grdpy");
            oGrid = ((SAPbouiCOM.Grid)(oitem.Specific));


            //crear el datasources  y despues dentro de esto se ejecuta la query
          //  oForm.DataSources.DataTables.item("grPagdt");
            oGrid.DataTable = oForm.DataSources.DataTables.Item("grPagdt");
            oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT distinct t0.PayMethCod AS Código,t0.Descript as Descripción,t0.Active as Activo from OPYM  as t0 left join  ocrd as t1  on t0.PayMethCod= t1.pymcode where t0.type='O' ");
            

            //cheackbox en el campo activo 
            oGrid.Columns.Item("Activo").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            //para agregar el link buton a la grilla
            oGrid.Columns.Item("Código").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
            SAPbouiCOM.EditTextColumn oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Código");
            oEdit.LinkedObjectType = "147";
            RowNumberGrid(oGrid);
           
          
    
       
        }


        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Item oitem;
        //public SAPbouiCOM.EventFilter oEvent;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.LinkedButton LinkedButton2;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.ComboBox ComboBox4;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.EditText EditText19;
        private SAPbouiCOM.Folder Folder6;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.ComboBox ComboBox5;
        private SAPbouiCOM.StaticText StaticText19;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText20;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.Button Button2;


        public void grBusVlr () {



            SAPbobsCOM.Recordset oRecordset1 = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));

            oitem = oForm.Items.Item("grdpy");
            oGrid = ((SAPbouiCOM.Grid)(oitem.Specific));
            string vlGrilla = null;

            string Sql10 = "select t0.PayMethCod as 'grcode',t1.Cardcode as 'cod' from OPYM as t0  left join  ocrd as t1 on t0.PayMethCod= t1.pymcode  where t1.cardcode='" + EditText0.Value.ToString() + "'";
            oRecordset1.DoQuery(Sql10);
            string Extraerdequery10 = oRecordset1.Fields.Item("grcode").Value.ToString();
            vlGrilla = Extraerdequery10;
            oForm.DataSources.DataTables.Item("grPagdt");
            oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT distinct t0.PayMethCod AS 'Código',t0.Descript as' Descripción',t0.Active as 'Activo'  from OPYM  as t0 left join  ocrd as t1  on t0.PayMethCod= t1.pymcode where t0.type='O' ");
            //oGrid.DataTable = oForm.DataSources.DataTables.Item("grPagdt2");
            oGrid.Columns.Item("Activo").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            //para agregar el link buton a la grilla
            oGrid.Columns.Item("Código").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
            SAPbouiCOM.EditTextColumn oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Código");
            oEdit.LinkedObjectType = "147";

            //oCBC.ValidValues.Add(oRec.Fields.Item(0).Value.ToString(), oRec.Fields.Item(1).Value.ToString());
            //algo parecido a lo que hace el boto de fijar metodo, lo hago aca para que el viejaSel no sea -1 y pueda sacar la linea en negrita
            //string valor = (string)oGrid.DataTable.GetValue("Código", 2); esto es para saber lo que lleva una linea en especifico
            for (int iRows = 0; iRows <= oGrid.Rows.Count - 1; iRows++)
            {
                if ((string)oGrid.DataTable.GetValue("Código", iRows) == vlGrilla)
                {
                    if (viejaSel != -1)
                        oGrid.CommonSetting.SetRowFontStyle(viejaSel, SAPbouiCOM.BoFontStyle.fs_Plain);

                    oGrid.CommonSetting.SetRowFontStyle(iRows + 1, SAPbouiCOM.BoFontStyle.fs_Bold);
                    viejaSel = iRows + 1;

                }
          
               
            }
         
    
            // Extraerdequery1 = Convert.ToString(Grid2.Columns.Item("Descripción"));



            //}


        }



        //clickafter   para llenar el combobox del banco segun el país
        //private void ComboBox4_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //      CleanComboBox(ComboBox1);
        //    string SqlCad3 = ("select odsc.BankCode,odsc.BankName from odsc inner join ocry on odsc.Countrycod = ocry.code WHERE odsc.CountryCod='" + ComboBox4.Value.ToString() + "'");
        //    LoadComboQueryRecordset(SqlCad3, ComboBox1, "BankCode", "BankName", oCompany);
         
        //}

        private void ComboBox4_ComboSelectBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CleanComboBox(ComboBox1);
            string SqlCad3 = ("select odsc.BankCode,odsc.BankName from odsc inner join ocry on odsc.Countrycod = ocry.code WHERE odsc.CountryCod='" + ComboBox4.Value.ToString() + "'");
            LoadComboQueryRecordset(SqlCad3, ComboBox1, "BankCode", "BankName", oCompany);
            //combobox de cuentas de los bancos
           
        }
     
  
        private void ComboBox1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CleanComboBox(ComboBox3);
            string SqlCad4 = "select distinct DSC1.Account,DSC1.BankCode  from DSC1  inner join ODSC ON DSC1.BankCode=DSC1.BankCode  where DSC1.BankCode='" + ComboBox1.Value.ToString() + "'";
            // oApp.SetStatusBarMessage("El dato es " + ComboBox1.Value.ToString());
            LoadComboQueryRecordset(SqlCad4, ComboBox3, "Account", "", oCompany);
           

            // Folder0.Item.Click(); 
        }
        private void ComboBox3_ComboSelectBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

    

            //utilizamos los recorset mas que nada para recorrer la query que llena los textbox
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.ComboBox oComboApproval = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbctbc").Specific;
            if (oComboApproval.Value.Trim() != "")
            {

                SAPbouiCOM.EditText oEditStatus = (SAPbouiCOM.EditText)oForm.Items.Item("txtsucur").Specific;
                SAPbouiCOM.EditText oEditStatus2 = (SAPbouiCOM.EditText)oForm.Items.Item("txtbics").Specific;
                SAPbouiCOM.EditText oEditStatus3 = (SAPbouiCOM.EditText)oForm.Items.Item("txtcotbn").Specific;
                string SqlCad5 = "select distinct DSC1.BRANCH as sucursal , DSC1.SwiftNum as Swift,ControlKey as control   from DSC1 inner join ODSC ON DSC1.BankCode=ODSC.BankCode  where DSC1.BankCode=" + ComboBox1.Value.ToString() + "";
                //   oApp.SetStatusBarMessage("El dato es " + SqlCad5);
                oRecordset.DoQuery(SqlCad5);
                string Extraerdequery = oRecordset.Fields.Item("sucursal").Value.ToString();
                string Extraerdequery2 = oRecordset.Fields.Item("Swift").Value.ToString();
                string Extraerdequery3 = oRecordset.Fields.Item("control").Value.ToString();
                oEditStatus.Value = Extraerdequery;
                oEditStatus2.Value = Extraerdequery2;
                oEditStatus3.Value = Extraerdequery3;

                //cargamos el IBAN
                CleanComboBox(ComboBox5);
                string SqlCad6 = "select distinct DSC1.IBAN  from DSC1  inner join ODSC ON DSC1.BankCode=DSC1.BankCode  where DSC1.BankCode='" + ComboBox1.Value.ToString() + "'";
                // oApp.SetStatusBarMessage("El dato es " + ComboBox3.Value.ToString());
                LoadComboQueryRecordset(SqlCad6, ComboBox5, "IBAN", "", oCompany);
            }
        }

        //para descamarcar una opcion anterior en los metodos de pagos - grilla debes esta en modo ms_single
        int viejaSel = -1;
        String sValorGrid;
        public void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent )
        {
            BubbleEvent = true;
            try
            {

                if (oGrid.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
                {

                    if (viejaSel != -1)
                        oGrid.CommonSetting.SetRowFontStyle(viejaSel, SAPbouiCOM.BoFontStyle.fs_Plain);

                    int nSelecRow = (oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
                    viejaSel = nSelecRow + 1;
                    //oApp.SetStatusBarMessage("El dato es " + oldSelection + nSelecRow);
                    //seleccion  el row y se destaca en negrita
                    oGrid.CommonSetting.SetRowFontStyle(nSelecRow + 1, SAPbouiCOM.BoFontStyle.fs_Bold);
                     sValorGrid = Convert.ToString(oGrid.DataTable.GetValue("Código", nSelecRow));

                    // Button0.Caption = "Update";
                    //if (Button0.Caption=="OK"){

                        Button0.Caption ="Update";

                    //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                      
                    //}

                 
                }

            }
            catch (Exception)
            {

                oApp.SetStatusBarMessage("Error  en el boton:  " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);



            }


        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;
            SAPbobsCOM.BusinessPartners sboBP = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

            switch (Button0.Caption)
            {
                case "Add":
                    
                    // oForm = (SAPbouiCOM.Form)oApp.Forms.ActiveForm;



                    sboBP.CardCode = EditText0.Value.ToString();
                    sboBP.CardName = EditText1.Value.ToString();
                    sboBP.FederalTaxID = EditText2.Value.ToString();
                    sboBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
                    sboBP.CardForeignName = EditText3.Value.ToString();
                    sboBP.Address = EditText4.Value.ToString();
                    sboBP.City = EditText5.Value.ToString();
                    sboBP.ZipCode = EditText6.Value.ToString();
                    sboBP.County = EditText8.Value.ToString();
                    sboBP.Phone1 = EditText9.Value.ToString();
                    sboBP.Phone2 = EditText10.Value.ToString();
                    sboBP.Fax = EditText11.Value.ToString();
                    sboBP.EmailAddress = EditText12.Value.ToString();
                    sboBP.Country = ComboBox2.Value.ToString();
                    //creo esta variable  y tomo el valor seleccionado, ya que al convetirlo a string da problemas con los espacios
                    string vlSelec = ComboBox0.Selected.Value;
                    sboBP.PayTermsGrpCode = Convert.ToInt16(vlSelec);
                    sboBP.HouseBankCountry = ComboBox4.Value.ToString();
                    sboBP.HouseBank = ComboBox1.Value.ToString();
                    sboBP.HouseBankAccount = ComboBox3.Value.ToString();
                    sboBP.HouseBankBranch = EditText18.Value.ToString();

                    //PeymentMethodCode y BPPaymentMethods , deben ir juntos ya que sap te pide un valor por defecto, que se guardan en la ocrd y crd2
                    sboBP.PeymentMethodCode = sValorGrid;
                    sboBP.BPPaymentMethods.PaymentMethodCode = sValorGrid;


                    //sboBP.Add();

                    int iRetVal = sboBP.Add();

                    // cuando viene un 1 significa que viene con error 
                    if (iRetVal != 0)
                    {
                        //eso muestra el ultimo error al usuario en la barra de abajo de sap
                        oApp.SetStatusBarMessage("Error creando al Colaborador: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }
                    else
                    {

                        oApp.SetStatusBarMessage("Registro exitoso: " + EditText0.Value.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        //Application.SBO_Application.ActivateMenuItem("1288");

                        
                    }
                    break;
                case "Update":
                  
                    if (sboBP.GetByKey(EditText0.Value.ToString()))
                    
                    {
                        sboBP.CardName = EditText1.Value.ToString();
                        sboBP.FederalTaxID = EditText2.Value.ToString();
                        sboBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
                        sboBP.CardForeignName = EditText3.Value.ToString();
                        sboBP.Address = EditText4.Value.ToString();
                        sboBP.City = EditText5.Value.ToString();
                        sboBP.ZipCode = EditText6.Value.ToString();
                        sboBP.County = EditText8.Value.ToString();
                        sboBP.Phone1 = EditText9.Value.ToString();
                        sboBP.Phone2 = EditText10.Value.ToString();
                        sboBP.Fax = EditText11.Value.ToString();
                        sboBP.EmailAddress = EditText12.Value.ToString();
                        sboBP.Country = ComboBox2.Value.ToString();
                        //creo esta variable  y tomo el valor seleccionado, ya que al convetirlo a string da problemas con los espacios
                        string vlSelec1 = ComboBox0.Selected.Value;
                        sboBP.PayTermsGrpCode = Convert.ToInt16(vlSelec1);
                        sboBP.HouseBankCountry = ComboBox4.Value.ToString();
                        sboBP.HouseBank = ComboBox1.Value.ToString();
                        sboBP.HouseBankAccount = ComboBox3.Value.ToString();
                        sboBP.HouseBankBranch = EditText18.Value.ToString();
                        sboBP.PeymentMethodCode = sValorGrid;
                        sboBP.BPPaymentMethods.PaymentMethodCode = sValorGrid;
                        sboBP.FreeText = EditText19.Value.ToString();
                        int x = sboBP.Update();




                        if (x != 0)
                        {
                            //eso muestra el ultimo error al usuario en la barra de abajo de sap
                            oApp.SetStatusBarMessage("Error actualizando al Colaborador: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }
                        else
                        {

                            oApp.SetStatusBarMessage("Actualización exitosa de : " + EditText0.Value.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            //Application.SBO_Application.ActivateMenuItem("1288");
                            ////goto case "OK";
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE; 
                        }
                    }

                    break;
                case "Find":
                    //if (Button0.Caption == "Find")
                    //{
                    if (EditText0.Value.ToString() != "")
                    {
                        grBusVlr();
                    }
                  

                    break;
   
                case "OK":
                   // oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                 
                   //  Muestra el formulario

                    
                    

                    break;
            }
        }

        //creo un evento clickafter para cuando se cambie de buscar a agregar , asi no borra los datos que hay en la grilla
       
        private void EditText0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Button0.Caption == "Add" )
            {

                grillametpag();

                 

            }
            //else if (Button0.Caption == "Find" && lin1 != -1)
            //{
            //grBusVlr();
            //      lin1 = -1;
            //}
          

        }
        //UTILIZO EL CLICKAFTER EN EL FOLDER PARA RE BUSCAR POR EN EDITEXT 0 EL CARDCODE Y ASI PODER MARCAR CON NEGRITO LA LINEA DE METODO DE PAGO
        //ademas creo una variable lin que me permita sacar la linea negrita, cuando se seleccione otra y no se vuelva a pasar el if
       
        
        private void Folder1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            //string saberCola = "select Cardcode as 'cod' from ocrd  where cardcode='" + EditText0.Value.ToString() + "'";
            //oRecordset.DoQuery(saberCola);
            //string colaborador = oRecordset.Fields.Item("cod").Value.ToString();
            BubbleEvent = true;
    
           
            if (Button0.Caption == "OK")
                 {
                   
                         grBusVlr();

                 //}
        
             }
       
             //Button0.Caption = "Update";

        }


        //SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;
        //SAPbobsCOM.BusinessPartners sboBP = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
        //if (sboBP.GetByKey(EditText0.Value.ToString())) {
        //int elim=    sboBP.Remove();
        //    if (elim != 0)
        //    {
        //        //eso muestra el ultimo error al usuario en la barra de abajo de sap
        //        oApp.SetStatusBarMessage("Error  eliminador al Colaborador: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
        //    }
        //    else
        //    {

        //        oApp.SetStatusBarMessage("Eliminación exitosa: " + EditText0.Value.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
        //        //Application.SBO_Application.ActivateMenuItem("1288");


        //    }

        //}



       

     
    

    

  

     

      

            
    }
}