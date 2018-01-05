using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.ediciongrso", "ediciongrso.b1f")]
    class ediciongrso : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        //public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        //public SAPbouiCOM.Item Oitem;
        public ediciongrso()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txtedco").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txtednm").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txtedpr1").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmed").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txtedpr2").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           

        }

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
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();


             CleanComboBox(ComboBox0);
             string SqlCad = ("  SELECT CAST (U_CentroPyto AS VARCHAR) + '' + CAST (U_DeptoPyto AS VARCHAR) + '' + CAST (U_CodigoPyto AS VARCHAR) As Codigo,U_NombrePyto FROM [@PROYECTOSCOSTE]");
           // oApp.SetStatusBarMessage("El dato es " + SqlCad );
             LoadComboQueryRecordset(SqlCad, ComboBox0, "Codigo", "U_NombrePyto", oCompany);


        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;
              switch (oForm.Mode)
            {
            case SAPbouiCOM.BoFormMode.fm_ADD_MODE:

            oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
            oUserTable.Code = EditText0.Value.ToString();
            oUserTable.Name = EditText0.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = EditText0.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_NombreEDC").Value = EditText1.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_ProyectoEDC").Value = EditText2.Value.ToString();
            oUserTable.Add();

            int i = oUserTable.Add();
            if (i != 0)
            {
                oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
            {
                oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                oForm = oApp.Forms.Item("edm");
                oForm.Close();



            }

            break;
                case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:

                      oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
            //hago un getbykey para obtener el valor key necesario para actualizar los datos
            if (oUserTable.GetByKey(EditText0.Value.ToString()))
                // oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = EditText0.Value.ToString();
                oUserTable.UserFields.Fields.Item("U_NombreEDC").Value = EditText1.Value.ToString();
                oUserTable.UserFields.Fields.Item("U_ProyectoEDC").Value = EditText2.Value.ToString();

            //oUserTable.Update();


            int J = oUserTable.Update();

            //oApp.SetStatusBarMessage("valor"+ i);
            if (J != 0)
            {
                oApp.SetStatusBarMessage("Error en la actualización: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
            {
                oApp.SetStatusBarMessage("Exito en la actualizacón", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                //oForm = oApp.Forms.Item("fmacted");
                //oForm.Close();



            }


            break;
                case SAPbouiCOM.BoFormMode.fm_FIND_MODE:



            break;
            }
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText3;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.ComboBox oComboAprueba = ComboBox0;
            //  oApp.SetStatusBarMessage("El dato es " + oComboAprueba.Value.ToString());
            if (oComboAprueba.Value.Trim() != "")
            {
                SAPbouiCOM.EditText oEditStatus = EditText2;
                SAPbouiCOM.EditText oEditStatus2 = EditText3;

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
    }
}
