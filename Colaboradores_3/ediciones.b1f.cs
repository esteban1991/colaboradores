using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{


    [FormAttribute("Colaboradores_3.ediciones", "ediciones.b1f")]
    class ediciones : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Item oitem;
        public SAPbobsCOM.UserTable oUserTable;
        public ediciones()
        {
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Gr_edi").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("3").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("4").Specific));
            this.Button3.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button3_ClickBefore);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("5").Specific));
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
        public static void RowNumberGrid(SAPbouiCOM.Grid oGrid)
        {
            SAPbouiCOM.RowHeaders oHeader = null;
            oHeader = oGrid.RowHeaders;

            for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
            {
                oHeader.SetText(i, Convert.ToString(i + 1));

            }
        }
        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            RowNumberGrid(Grid0);

                    }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            cread_edi crear_form = new cread_edi();
            crear_form.Show();
        }

        private SAPbouiCOM.Button Button2;

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
          
            oForm = oApp.Forms.Item("edform");
            //  Muestra el formulario
            //oForm.Visible = true;
            //busca la grilla con el UID
           
            oitem = oForm.Items.Item("Gr_edi");
            //Grid0 = ((SAPbouiCOM.Grid)(oitem.Specific));
            // ejecuta la query
            oForm.DataSources.DataTables.Item(0).ExecuteQuery("select t0.U_CodigoEDC AS 'Código' ,t0.U_NombreEDC as 'Nombre',t0.U_ProyectoEDC as 'Proyecto',t1.U_NombrePyto as [Nombre Proyecto] from [@EDICIONESCOL] as t0   left join [@PROYECTOSCOSTE] as t1 on t0.U_ProyectoEDC = (CAST (t1.U_CentroPyto AS VARCHAR) + '' + CAST (t1.U_DeptoPyto AS VARCHAR) + '' + CAST (t1.U_CodigoPyto AS VARCHAR)) ");
            Grid0.DataTable = oForm.DataSources.DataTables.Item("edi_dt");
            RowNumberGrid(Grid0);
         
        }

        private SAPbouiCOM.Button Button3;

        private void Button3_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (Grid0.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
            {
                int nSelecRow = (Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
                String sValorGrid = Convert.ToString(Grid0.DataTable.GetValue("Código", nSelecRow));
               // oApp.SetStatusBarMessage("codigo" + sValorGrid);
                actua_edi crear_form = new actua_edi(sValorGrid);
                crear_form.Show();
            }
            else
            {
                oApp.SetStatusBarMessage("Error, no has seleccionado una fila"  , SAPbouiCOM.BoMessageTime.bmt_Medium, false);
            }
        }

        private SAPbouiCOM.Button Button4;

        private void Button4_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

         if (Grid0.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
         {
             int nSelecRow = (Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder));
             String sValorGrid = Convert.ToString(Grid0.DataTable.GetValue("Código", nSelecRow));
             oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
             if (oUserTable.GetByKey(sValorGrid)) {

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
                     Button2_ClickBefore(sboObject, pVal, out BubbleEvent);


                 }
             }

         }
         else
         {
             oApp.SetStatusBarMessage("Error, no has seleccionado una fila", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
         }
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
         

        }

      

   
    }
}
