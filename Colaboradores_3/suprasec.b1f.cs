using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.suprasec", "suprasec.b1f")]
    class suprasec : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        public suprasec()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GRSPSEC").Specific));
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.OnCustomInitialize();

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

        private void OnCustomInitialize()
        {

            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();

            oForm = oApp.Forms.Item("supra_sec");
         
            Grid0.Columns.Item("Código").TitleObject.Sortable = true;
            Grid0.DataTable.Rows.Add(1);

            for (int i = 1; i <= this.Grid0.DataTable.Rows.Count; i += 1)
            {

                if (i < this.Grid0.DataTable.Rows.Count)
                {

                    Grid0.Rows.SelectedRows.Add(i);
                }

            }
            RowNumberGrid(Grid0);
        }

        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oUserTable = oCompany.UserTables.Item("SUPRASECCIONESCOL");
            if (oForm.Mode.Equals(SAPbouiCOM.BoFormMode.fm_OK_MODE) || (Grid0.Rows.SelectedRows.Count == 0))
            {


                Button0.Caption = "OK";

            }

            else
            {

                //if (Grid2.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
                //{


                int nRow = Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                oApp.SendKeys("({TAB})");//aca aplico tabular para que tome el campo para actualizar




                //Grid2.Rows.SelectedRows.Equals(nRow);
                String sValorGrid = Convert.ToString(Grid0.DataTable.GetValue("Código", nRow));

                // Grid2.Columns.Item("CODE").Click();
                //bool num = (Grid2.Rows.SelectedRows.Count > 0);
                if (oUserTable.GetByKey(sValorGrid.ToString())) // Esto devuelve true si existe el registro
                {

                    //

                    string nom = (string)(Grid0.DataTable.GetValue("Nombre", nRow));




                    // oCompany.StartTransaction();
                    oUserTable.Code = sValorGrid;
                    oUserTable.Name = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_CodigoSS").Value = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_NombreSS").Value = nom;


                    int i = oUserTable.Update();

                    if (i != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la Actualización", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL]");
                        Grid0.DataTable = oForm.DataSources.DataTables.Item("DTSPSEC");

                        Grid0.DataTable.Rows.Add(1);
                        for (int j = 1; j <= this.Grid0.DataTable.Rows.Count; j += 1)
                        {

                            if (j < this.Grid0.DataTable.Rows.Count)
                            {

                                Grid0.Rows.SelectedRows.Add(j);
                            }

                        }

                        RowNumberGrid(Grid0);
                        BubbleEvent = false;

                    }


                }


                     //si no existe el dato, lo agregara
                else
                {


                    //   Button0.Caption = "Agregar";
                    //   oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
                    int nRow2 = Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    String sValorGrid2 = Convert.ToString(Grid0.DataTable.GetValue("Código", nRow2));
                    //string nNOM2 = (string)Grid2.DataTable.GetValue("NOM", nRow2);
                    // string cod2 = (string)(Grid2.DataTable.GetValue("Código", nRow2));
                    string nom2 = (string)(Grid0.DataTable.GetValue("Nombre", nRow2));


                    oUserTable.Code = sValorGrid2;
                    oUserTable.Name = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_CodigoSS").Value = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_NombreSS").Value = nom2;

                    int j = oUserTable.Add();

                    if (j != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                        Grid0.DataTable.Rows.Add(1);

                        oForm.DataSources.DataTables.Item(0).ExecuteQuery("SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL]");
                        Grid0.DataTable = oForm.DataSources.DataTables.Item("DTSPSEC");

                        for (int i = 1; i <= this.Grid0.DataTable.Rows.Count; i += 1)
                        {

                            if (i < this.Grid0.DataTable.Rows.Count)
                            {

                                Grid0.Rows.SelectedRows.Add(i);
                            }

                        }

                        RowNumberGrid(Grid0);
                        BubbleEvent = false;



                    }
                }

            }

        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid0.Rows.SelectedRows.Clear();

            Grid0.Rows.SelectedRows.Add(pVal.Row);
            


        }
    }
}
