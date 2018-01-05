using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.actua_secc", "actua_secc.b1f")]
    class actua_secc : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        private string sValorGrid2;

        public actua_secc()
        {
        }

        public actua_secc(string sValorGrid)
        {
            // TODO: Complete member initialization
            this.sValorGrid2 = sValorGrid;
            EditText0.Value = sValorGrid2;
            OnInitializeComponent();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txtcodse").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txtnosecc").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txtsp1").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmactsec").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txtsp2").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txtprase").Specific));
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
            //busca la conexion de la empresa actual
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();



            oForm = oApp.Forms.Item("actsecc");
            ////  Muestra el formulario
            oForm.Visible = true;

            //aca creo el data source manualmente , es decir desde el cuadro de herramientas
           
            oForm.DataSources.UserDataSources.Item("actprec");
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("txtprase").Specific;
            oEdit.DataBind.SetBound(true, "", "actprec");
            
            //busca con recorset los datos y los muestros en los textboxes
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.EditText oEditStatus = EditText1;
            SAPbouiCOM.EditText oEditStatus2 = EditText2;
            SAPbouiCOM.EditText oEditStatus3 = EditText3;
            SAPbouiCOM.EditText oEditStatus4 = EditText4;
            //oApp.SetStatusBarMessage("El dato es..." + EditText0.Value);
            string SqlCad1 = "SELECT T0.U_NombreSEC AS 'NOMBRE',T0.U_SupraSeccionSEC as 'codsupra',T1.U_NombreSS AS [SUPRA-SECCIÓN] ,T0.U_PrecioSEC AS 'PRECIO' FROM [@SECCIONESCOL] AS T0 LEFT JOIN [@SUPRASECCIONESCOL] AS T1 ON T0.U_SupraSeccionSEC=T1.U_CodigoSS where T0.U_CodigoSEC='" + sValorGrid2 + "'";
            // oApp.SetStatusBarMessage("El dato es " + SqlCad1);
            oRecordset.DoQuery(SqlCad1);
            string Extraerdequery = oRecordset.Fields.Item("NOMBRE").Value.ToString();
            string Extraerdequery2 = oRecordset.Fields.Item("codsupra").Value.ToString();
            string Extraerdequery3 = oRecordset.Fields.Item("SUPRA-SECCIÓN").Value.ToString();
            string Extraerdequery4 = oRecordset.Fields.Item("PRECIO").Value.ToString();
            decimal precio = Convert.ToDecimal(Extraerdequery4);
            //oApp.SetStatusBarMessage("El dato es " + precio );
        
            oEditStatus.Value = Extraerdequery;
            oEditStatus2.Value = Extraerdequery2;
            oEditStatus3.Value = Extraerdequery3;
            oEditStatus4.Value = Extraerdequery4;


            CleanComboBox(ComboBox0);
            string SqlCad = ("SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL]");
            // oApp.SetStatusBarMessage("El dato es " + SqlCad );
            LoadComboQueryRecordset(SqlCad, ComboBox0, "Código", "Nombre", oCompany);


        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText3;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
       

            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            SAPbouiCOM.ComboBox oComboAprueba = ComboBox0;
            //  oApp.SetStatusBarMessage("El dato es " + oComboAprueba.Value.ToString());
            if (oComboAprueba.Value.Trim() != "")
            {
                SAPbouiCOM.EditText oEditStatus3 = EditText2;
                SAPbouiCOM.EditText oEditStatus4 = EditText3;

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

        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText3;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;


            oUserTable = oCompany.UserTables.Item("SECCIONESCOL");
            //hago un getbykey para obtener el valor key necesario para actualizar los datos
            if (oUserTable.GetByKey(EditText0.Value.ToString()))
                // oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = EditText0.Value.ToString();
                oUserTable.UserFields.Fields.Item("U_NombreSEC").Value = EditText1.Value.ToString();
                oUserTable.UserFields.Fields.Item("U_SupraSeccionSEC").Value = EditText2.Value.ToString();
                oUserTable.UserFields.Fields.Item("U_PrecioSEC").Value = EditText4.Value.ToString();
            

            //oUserTable.Update();


            int i = oUserTable.Update();

            //oApp.SetStatusBarMessage("valor"+ i);
            if (i != 0)
            {
                oApp.SetStatusBarMessage("Error en la actualización: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

            }
            else
            {
                oApp.SetStatusBarMessage("Exito en la actualizacón", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                //oForm = oApp.Forms.Item("fmacted");
                //oForm.Close();



            }

        }
    }
}
