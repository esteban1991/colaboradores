using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.cread_edi", "cread_edi.b1f")]
    class cread_edi : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        
        public cread_edi()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            //     oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            //     oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cb_Pred1").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txt_Cred").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("txt_Pred1").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("txt_Pred2").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("txt_nmed").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

      

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
        private SAPbouiCOM.EditText EditText4;
        //private SAPbouiCOM.StaticText StaticText1;
        //private SAPbouiCOM.EditText EditText1;
        //private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.Button Button2;
        //private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText7;
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




        private void OnCustomInitialize()
        {

             CleanComboBox(ComboBox0);
             string SqlCad = ("  SELECT CAST (U_CentroPyto AS VARCHAR) + '' + CAST (U_DeptoPyto AS VARCHAR) + '' + CAST (U_CodigoPyto AS VARCHAR) As Codigo,U_NombrePyto FROM [@PROYECTOSCOSTE]");
           // oApp.SetStatusBarMessage("El dato es " + SqlCad );
             LoadComboQueryRecordset(SqlCad, ComboBox0, "Codigo", "U_NombrePyto", oCompany);

            
        }
        


        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            

        }


        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.ComboBox oComboAprueba = ComboBox0;
            //  oApp.SetStatusBarMessage("El dato es " + oComboAprueba.Value.ToString());
            if (oComboAprueba.Value.Trim() != "")
            {
                SAPbouiCOM.EditText oEditStatus = EditText5;
                SAPbouiCOM.EditText oEditStatus2 = EditText6;

                string SqlCad1 = "SELECT (CAST (U_CentroPyto AS VARCHAR) + '' + CAST (U_DeptoPyto AS VARCHAR) + '' + CAST (U_CodigoPyto AS VARCHAR)) As Code ,U_NombrePyto FROM [@PROYECTOSCOSTE] where (CAST (U_CentroPyto AS VARCHAR) + '' + CAST (U_DeptoPyto AS VARCHAR) + '' + CAST (U_CodigoPyto AS VARCHAR))=" + ComboBox0.Value.ToString() + "";
                oRecordset.DoQuery(SqlCad1);
              // oApp.SetStatusBarMessage("El dato es " + SqlCad1 );
                string Extraerdequery = oRecordset.Fields.Item("Code").Value.ToString();
                string Extraerdequery2 = oRecordset.Fields.Item("U_NombrePyto").Value.ToString();
                //oApp.SetStatusBarMessage("El dato es " + Extraerdequery+"Y EL DOS" +Extraerdequery2);
                oEditStatus.Value = Extraerdequery;
                oEditStatus2.Value = Extraerdequery2;

            }


        }


        
      

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            //oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();

            oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
            oUserTable.Code = EditText4.Value.ToString();
            oUserTable.Name = EditText4.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = EditText4.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_NombreEDC").Value = EditText7.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_ProyectoEDC").Value = EditText5.Value.ToString();
            //oUserTable.Add();
          
            int i = oUserTable.Add();
            

            if (i != 0)
            {
                oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                
            }
            else
            {
                oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                oForm = oApp.Forms.Item("fmcred");
                oForm.Close();
                
               
                
            }
        }

       



       
       

  
       

       

  

    

      



       
    }
}
