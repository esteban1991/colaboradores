using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.Secciones", "Secciones.b1f")]
    class Secciones : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Item oitem;
        public SAPbobsCOM.UserTable oUserTable;
        public Secciones()
        {

            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GR_SECC").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("txtactsec").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("txtrefsec").Specific));
            this.Button3.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button3_ClickBefore);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("txtelmsec").Specific));
            this.Button4.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button4_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private SAPbouiCOM.Button Button0;

        public static void RowNumberGrid(SAPbouiCOM.Grid oGrid)
        {
            SAPbouiCOM.RowHeaders oHeader = null;
            oHeader = oGrid.RowHeaders;

            for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
            {
                oHeader.SetText(i, Convert.ToString(i + 1));

            }
        }



        private void OnCustomInitialize()
        {


            RowNumberGrid(Grid0);

        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {


        }

        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (Grid0.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
            {
                int nSelecRow = (Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
                String sValorGrid = Convert.ToString(Grid0.DataTable.GetValue("CÓDIGO", nSelecRow));
                // oApp.SetStatusBarMessage("codigo" + sValorGrid);
                actua_secc crear_form = new actua_secc(sValorGrid);
                crear_form.Show();
            }
            else
            {
                oApp.SetStatusBarMessage("Error, no has seleccionado una fila", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }

        private void Button3_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            oForm = oApp.Forms.Item("seccform");
            //  Muestra el formulario
            //oForm.Visible = true;
            //busca la grilla con el UID

            oitem = oForm.Items.Item("GR_SECC");
            //Grid0 = ((SAPbouiCOM.Grid)(oitem.Specific));
            // ejecuta la query
            oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.U_CodigoSEC AS 'CÓDIGO',T0.U_NombreSEC AS 'NOMBRE',T1.U_NombreSS AS [SUPRA-SECCIÓN] ,T0.U_PrecioSEC AS 'PRECIO' FROM [@SECCIONESCOL] AS T0 LEFT JOIN [@SUPRASECCIONESCOL] AS T1 ON T0.U_SupraSeccionSEC=T1.U_CodigoSS ");
            Grid0.DataTable = oForm.DataSources.DataTables.Item("DT_SECC");
            RowNumberGrid(Grid0);
        }

        private void Button4_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (Grid0.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
            {
                int nSelecRow = (Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
                String sValorGrid = Convert.ToString(Grid0.DataTable.GetValue("CÓDIGO", nSelecRow));
                oUserTable = oCompany.UserTables.Item("SECCIONESCOL");
                if (oUserTable.GetByKey(sValorGrid))
                {

                    int i = oUserTable.Remove();


                    if (i != 0)
                    {
                        oApp.SetStatusBarMessage("Error al eliminar : " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Edición Eliminada", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        //oForm = oApp.Forms.Item("fmacted");
                        //oForm.Close();
                         //presiono el boton refrescar   
                        Button3_ClickBefore(sboObject, pVal, out BubbleEvent);

                    }
                }

            }
            else
            {
                oApp.SetStatusBarMessage("Error, no has seleccionado una fila", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            crea_secc crear_form = new crea_secc();
            crear_form.Show();
        }
    }
}
