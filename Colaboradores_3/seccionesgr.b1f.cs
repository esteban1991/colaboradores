using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.seccionesgr", "seccionesgr.b1f")]
    class seccionesgr : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        public seccionesgr()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GRSEC").Specific));
            this.Grid0.ComboSelectAfter += new SAPbouiCOM._IGridEvents_ComboSelectAfterEventHandler(this.Grid0_ComboSelectAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
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
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
           // Grid0.Columns.Item("CODSUPRA").Visible = false;
            oForm = oApp.Forms.Item("secci");
            Grid0.Columns.Item("Código").TitleObject.Sortable = true;
            Grid0.DataTable.Rows.Add(1);
            gridreco();
            //selecciona la ultima linea 
            for (int i = 1; i <= this.Grid0.DataTable.Rows.Count; i += 1)
            {

                if (i < this.Grid0.DataTable.Rows.Count)
                {

                    Grid0.Rows.SelectedRows.Add(i);
                }

            }
            gridset();
            //numerar la grilla
            RowNumberGrid(Grid0);
        }
     
        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            

        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid0.Rows.SelectedRows.Clear();
            
            Grid0.Rows.SelectedRows.Add(pVal.Row);

        }

        public void gridset() {

            //desahabilitar un campo de la grilla 
            SAPbouiCOM.CommonSetting setear = Grid0.CommonSetting;
            Grid0.Columns.Item("Código").TitleObject.Sortable = true;

            for (int i = 1; i <= this.Grid0.DataTable.Rows.Count; i += 1)
            {


                setear.SetCellEditable(i, 4, false);

            }
        }


        public void gridreco()
        {
            //combobox en en un grid
        
            Grid0.Columns.Item("CODSUPRA").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
            SAPbouiCOM.ComboBoxColumn oCBC = (SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("CODSUPRA");
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
         
            oRec.DoQuery("SELECT U_CodigoSS AS 'Código',U_NombreSS AS 'Nombre' FROM  [@SUPRASECCIONESCOL]");
            
            oRec.MoveFirst();
            while (!oRec.EoF)
            {
                oCBC.ValidValues.Add(oRec.Fields.Item(0).Value.ToString(), oRec.Fields.Item(1).Value.ToString());
                //mostrar la descripción en el combobox
             //   oCBC.DisplayType = (SAPbouiCOM.BoComboDisplayType.cdt_Description);
               // string Extraerdequery1 = oRec.Fields.Item("Código").Value.ToString();
                string Extraerdequery = oRec.Fields.Item("Nombre").Value.ToString();
               // Grid0.Columns.Item("SUPRA-SECCIÓN").T = oRec.Fields.Item("Nombre").Value.ToString();
                //(Grid0.Columns.Item("SUPRA-SECCIÓN")).(Extraerdequery);
                oRec.MoveNext();
                
            }
        }
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            oUserTable = oCompany.UserTables.Item("SECCIONESCOL");


                        

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
                    string suprse = (string)(Grid0.DataTable.GetValue("CODSUPRA", nRow));
                    double prec = (double)(Grid0.DataTable.GetValue("Precio", nRow));




                    // oCompany.StartTransaction();
                    oUserTable.Code = sValorGrid;
                    oUserTable.Name = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_CodigoSEC").Value = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_NombreSEC").Value = nom;
                    oUserTable.UserFields.Fields.Item("U_SupraSeccionSEC").Value = suprse;
                    oUserTable.UserFields.Fields.Item("U_PrecioSEC").Value = prec;

                    int i = oUserTable.Update();

                    if (i != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la Actualización", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        oForm.DataSources.DataTables.Item(0).ExecuteQuery(" SELECT T0.U_CodigoSEC AS 'Código',T0.U_NombreSEC AS 'Nombre',t0.U_SupraSeccionSEC AS 'CODSUPRA',T1.U_NombreSS AS [SUPRA-SECCIÓN],T0.U_PrecioSEC AS 'Precio' FROM [@SECCIONESCOL] AS T0 LEFT JOIN [@SUPRASECCIONESCOL] AS T1 ON T0.U_SupraSeccionSEC=T1.U_CodigoSS ");
                        Grid0.DataTable = oForm.DataSources.DataTables.Item("DTSEC");
                      //  Grid0.Columns.Item("CODSUPRA").Visible = false;
                      
                        for (int j = 1; j <= this.Grid0.DataTable.Rows.Count; j += 1)
                        {

                            if (j < this.Grid0.DataTable.Rows.Count)
                            {

                                Grid0.Rows.SelectedRows.Add(j);
                            }

                        }
                        Grid0.DataTable.Rows.Add(1);
                        gridreco();
                        gridset();
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
                    string suprse2 = (string)(Grid0.DataTable.GetValue("CODSUPRA", nRow));
                    double prec2 = (double)(Grid0.DataTable.GetValue("Precio", nRow));

                    oUserTable.Code = sValorGrid2;
                    oUserTable.Name = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_CodigoSEC").Value = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_NombreSEC").Value = nom2;
                    oUserTable.UserFields.Fields.Item("U_SupraSeccionSEC").Value = suprse2;
                    oUserTable.UserFields.Fields.Item("U_PrecioSEC").Value = prec2;
                    int j = oUserTable.Add();

                    if (j != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                      
                        oForm.DataSources.DataTables.Item(0).ExecuteQuery(" SELECT T0.U_CodigoSEC AS 'Código',T0.U_NombreSEC AS 'Nombre',t0.U_SupraSeccionSEC AS 'CODSUPRA',T1.U_NombreSS AS [SUPRA-SECCIÓN],T0.U_PrecioSEC AS 'Precio' FROM [@SECCIONESCOL] AS T0 LEFT JOIN [@SUPRASECCIONESCOL] AS T1 ON T0.U_SupraSeccionSEC=T1.U_CodigoSS ");
                        Grid0.DataTable = oForm.DataSources.DataTables.Item("DTSEC");
                        Grid0.DataTable.Rows.Add(1);
                        gridreco();

                    //    Grid0.Columns.Item("CODSUPRA").Visible = false;
                        for (int i = 1; i <= this.Grid0.DataTable.Rows.Count; i += 1)
                        {

                            if (i < this.Grid0.DataTable.Rows.Count)
                            {

                                Grid0.Rows.SelectedRows.Add(i);
                            }

                        }
                        gridset();
                        RowNumberGrid(Grid0);
                        BubbleEvent = false;



                    }
                }

            }







        }

        private void Grid0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if ((pVal.ColUID == "CODSUPRA"))
            {

                int dtRow = Grid0.GetDataTableRowIndex(pVal.Row);
                SAPbouiCOM.ComboBoxColumn oCBC = (SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("CODSUPRA");
             
                string ValSelec = oCBC.GetSelectedValue(pVal.Row).Value; // Selecciona El Valor (U_CodigoSS)
                string DesSelec = oCBC.GetSelectedValue(pVal.Row).Description; // Selecciona la Descripcion (U_NombreSS)
                //mostrar la descripción en el combobox
              //  oCBC.DisplayType = (SAPbouiCOM.BoComboDisplayType.cdt_Description);
                // Para Asignar el Valor a una celda del grid se puede asi:
                Grid0.DataTable.SetValue("SUPRA-SECCIÓN", dtRow, DesSelec);
               
             
                //O tambien asi
                //Grid0.DataTable.Columns.Item("CODSUPRA").Cells(pVal.Row).value = ValSelec;

            }

        }
    }
}
