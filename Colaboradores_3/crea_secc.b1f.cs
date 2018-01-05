using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.crea_secc", "crea_secc.b1f")]
    class crea_secc : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        //public SAPbouiCOM.Item Oitem;
        public crea_secc()
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
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txtcodsec").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txtnoseca").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txtprsec").Specific));
       
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txtspa1").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txtspa2").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmcresec").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
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
            //conexion  actual
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();


            CleanComboBox(ComboBox0);
            string SqlCad = ("SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL]");
            // oApp.SetStatusBarMessage("El dato es " + SqlCad );
            LoadComboQueryRecordset(SqlCad, ComboBox0, "Código", "Nombre", oCompany);



            //activamos el form para poder dar formato de moenda al campo precio
            oForm = oApp.Forms.Item("creasecc");
            //  Muestra el formulario
            oForm.Visible = true;
            //EditText2.Value = "0,00";
            //creamos un userdatasources , con estilo precio, luego se lo asignamos al campo precio (txtprsec)
            oForm.DataSources.UserDataSources.Add("VALORPRE", SAPbouiCOM.BoDataType.dt_PRICE, 0);
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("txtprsec").Specific;
            oEdit.DataBind.SetBound(true, "", "VALORPRE");
            //'Asignar Valor a EditText por Medio del UDS '
            SAPbouiCOM.UserDataSource oUserDataSource = oForm.DataSources.UserDataSources.Item("VALORPRE");
            oUserDataSource.ValueEx = "0.00";
           
          
        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText3;


        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.ComboBox oComboAprueba = ComboBox0;
            //  oApp.SetStatusBarMessage("El dato es " + oComboAprueba.Value.ToString());
            if (oComboAprueba.Value.Trim() != "")
            {
                SAPbouiCOM.EditText oEditStatus3 = EditText3;
                SAPbouiCOM.EditText oEditStatus4 = EditText4;

                string SqlCad1 = "SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL] where U_CodigoSS=" + ComboBox0.Value.ToString() + "";
                oRecordset.DoQuery(SqlCad1);
                // oApp.SetStatusBarMessage("El dato es " + SqlCad1 );
                string Extraerdequery = oRecordset.Fields.Item("Código").Value.ToString();
                string Extraerdequery2 = oRecordset.Fields.Item("Nombre").Value.ToString();
                //oApp.SetStatusBarMessage("El dato es " + Extraerdequery+"Y EL DOS" +Extraerdequery2);
                oEditStatus3.Value = Extraerdequery;
                oEditStatus4.Value = Extraerdequery2;
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            oUserTable = oCompany.UserTables.Item("SECCIONESCOL");
            oUserTable.Code = EditText0.Value.ToString();
            oUserTable.Name = EditText0.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_CodigoSEC").Value = EditText0.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_NombreSEC").Value = EditText1.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_SupraSeccionSEC").Value = EditText3.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_PrecioSEC").Value = EditText2.Value.ToString();
            oUserTable.UserFields.Fields.Item("U_ProyectoSEC").Value = "";
            //oUserTable.Add();

            int i = oUserTable.Add();


            if (i != 0)
            {
                oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
            {
                oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                oForm = oApp.Forms.Item("creasecc");
                oForm.Close();
                


            }
        }

      

        
    }
}
