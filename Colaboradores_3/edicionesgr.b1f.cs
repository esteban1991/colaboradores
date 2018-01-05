using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.edicionesgr", "edicionesgr.b1f")]
    class edicionesgr : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbobsCOM.UserTable oUserTable;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
       
        public edicionesgr()
        {
       
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            //                      this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("grilaedi").Specific));
            this.Grid2.ComboSelectAfter += new SAPbouiCOM._IGridEvents_ComboSelectAfterEventHandler(this.Grid2_ComboSelectAfter);
            this.Grid2.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid2_ClickAfter);
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


        public static void RowNumberGrid(SAPbouiCOM.Grid oGrid)
        {
            SAPbouiCOM.RowHeaders oHeader = null;
            oHeader = oGrid.RowHeaders;

            for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
            {
                oHeader.SetText(i, Convert.ToString(i + 1));

            }
        }


       //public String Col_StringDatoAnterior ;


        //private void Grid2_GotFocusAfter(System.Object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    if ((pVal.ColUID == "Código" || pVal.ColUID == "Nombre" || pVal.ColUID == "Proyecto"))
        //    {
        //        int dtRow = Grid2.GetDataTableRowIndex(pVal.Row);
        //        Col_StringDatoAnterior = Grid2.DataTable.GetValue(pVal.ColUID, dtRow).ToString(); //Asigno el valor inicial de la celda
        //    }
        //}


        //private void Grid2_ValidateBefore(System.Object sboObject, SAPbouiCOM.SBOItemEventArg pVal, ref System.Boolean BubbleEvent)
        //{
        //    oForm.Freeze(true);
        //    if (Grid2.DataTable.GetValue(pVal.ColUID, pVal.Row).ToString() != Col_StringDatoAnterior)
        //    {

        //        Button0.Caption = "Actualizar";


        //        oForm.Freeze(false);
        //    }
        //}

        private void OnCustomInitialize()
        {

            //conexion
            oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
           
            oForm = oApp.Forms.Item("grilaedi");
          //  OCULTAR CAMPOS QUE SIRVEN PARA REFERENCIAR;
         //    Grid2.Columns.Item("CODE").Visible = false;
            //Grid2.Columns.Item("NOM").Visible = false;
            //permite que se permita ordenar este campo
            Grid2.Columns.Item("Código").TitleObject.Sortable = true;
            //Grid2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            
            //AÑADO UNA LINEA VACIA A MI DATABLE 

            Grid2.DataTable.Rows.Add(1);
            gridreco();
          //  Grid2.Rows.SelectedRows.Add(i);
           // Grid2.GetCellFocus();

            for (int i = 1; i <= this.Grid2.DataTable.Rows.Count; i += 1)
            {

                if (i < this.Grid2.DataTable.Rows.Count) {

                Grid2.Rows.SelectedRows.Add(i);
                }

            }
            
            //numerar la grilla
           RowNumberGrid(Grid2);
           gridset();
         
                    
        }
        public void gridset() {
            
            //desahabilitar un campo de la grilla 
            SAPbouiCOM.CommonSetting setear = Grid2.CommonSetting;
            Grid2.Columns.Item("Código").TitleObject.Sortable = true;

            for (int i = 1; i <= this.Grid2.DataTable.Rows.Count; i += 1)
            {


                setear.SetCellEditable(i, 4, false);

            }
          
        
        
        }
        public void gridreco(){
            //combobox en en un grid
            Grid2.Columns.Item("Proyecto").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
            SAPbouiCOM.ComboBoxColumn oCBC = (SAPbouiCOM.ComboBoxColumn)Grid2.Columns.Item("Proyecto");
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(" SELECT CAST (U_CentroPyto AS VARCHAR) + '' + CAST (U_DeptoPyto AS VARCHAR) + '' + CAST (U_CodigoPyto AS VARCHAR) As Codigo,U_NombrePyto FROM [@PROYECTOSCOSTE]");

            oRec.MoveFirst();
            while (!oRec.EoF)
            {
                oCBC.ValidValues.Add(oRec.Fields.Item(0).Value.ToString(), oRec.Fields.Item(1).Value.ToString());
                string Extraerdequery1 = oRec.Fields.Item("U_NombrePyto").Value.ToString();

                Extraerdequery1 = Convert.ToString(Grid2.Columns.Item("Descripción"));
                
                oRec.MoveNext();
               
            }

          
            oCBC = null;
        
        }
      
        private SAPbouiCOM.Grid Grid2;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
           
            oUserTable = oCompany.UserTables.Item("EDICIONESCOL");

            


          //if (Grid2.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO)
          //          {

            if (oForm.Mode.Equals(SAPbouiCOM.BoFormMode.fm_OK_MODE) || (Grid2.Rows.SelectedRows.Count == 0))
            {


                Button0.Caption = "OK";
                
            }


            else
            {

                //if (Grid2.Rows.SelectedRows.Count > 0)   //VERIFICA QUE EXISTA UN ROW SELECCIONADO
                //{


                int nRow = Grid2.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                oApp.SendKeys("({TAB})");//aca aplico tabular para que tome el campo para actualizar




                //Grid2.Rows.SelectedRows.Equals(nRow);
                String sValorGrid = Convert.ToString(Grid2.DataTable.GetValue("Código", nRow));

                // Grid2.Columns.Item("CODE").Click();
                //bool num = (Grid2.Rows.SelectedRows.Count > 0);
                if (oUserTable.GetByKey(sValorGrid.ToString())) // Esto devuelve true si existe el registro
                {

                    //

                    string nom = (string)(Grid2.DataTable.GetValue("Nombre", nRow));
                    string proy = (string)(Grid2.DataTable.GetValue("Proyecto", nRow));




                    // oCompany.StartTransaction();
                    oUserTable.Code = sValorGrid;
                    oUserTable.Name = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = sValorGrid;
                    oUserTable.UserFields.Fields.Item("U_NombreEDC").Value = nom;
                    oUserTable.UserFields.Fields.Item("U_ProyectoEDC").Value = proy;

                    int i = oUserTable.Update();

                    if (i != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la Actualización", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        oForm.DataSources.DataTables.Item(0).ExecuteQuery("select  t0.U_CodigoEDC AS 'Código' ,t0.U_NombreEDC as 'Nombre',t0.U_ProyectoEDC as 'Proyecto',t1.U_NombrePyto as 'Descripción' from [@EDICIONESCOL] as t0   left join [@PROYECTOSCOSTE] as t1 on t0.U_ProyectoEDC = (CAST (t1.U_CentroPyto AS VARCHAR) + '' + CAST (t1.U_DeptoPyto AS VARCHAR) + '' + CAST (t1.U_CodigoPyto AS VARCHAR)) ");
                        Grid2.DataTable = oForm.DataSources.DataTables.Item("dted");

                        Grid2.DataTable.Rows.Add(1);
                        for (int j = 1; j <= this.Grid2.DataTable.Rows.Count; j += 1)
                        {

                            if (j < this.Grid2.DataTable.Rows.Count)
                            {

                                Grid2.Rows.SelectedRows.Add(j);
                            }

                        }
                        gridreco();
                        gridset();
                        RowNumberGrid(Grid2);
                        BubbleEvent = false;

                    }


                }


                     //si no existe el dato, lo agregara
                else
                {


                    //   Button0.Caption = "Agregar";
                    //   oUserTable = oCompany.UserTables.Item("EDICIONESCOL");
                    int nRow2 = Grid2.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    String sValorGrid2 = Convert.ToString(Grid2.DataTable.GetValue("Código", nRow2));
                    //string nNOM2 = (string)Grid2.DataTable.GetValue("NOM", nRow2);
                    // string cod2 = (string)(Grid2.DataTable.GetValue("Código", nRow2));
                    string nom2 = (string)(Grid2.DataTable.GetValue("Nombre", nRow2));
                    string proy2 = (string)(Grid2.DataTable.GetValue("Proyecto", nRow));

                    oUserTable.Code = sValorGrid2;
                    oUserTable.Name = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_CodigoEDC").Value = sValorGrid2;
                    oUserTable.UserFields.Fields.Item("U_NombreEDC").Value = nom2;
                    oUserTable.UserFields.Fields.Item("U_ProyectoEDC").Value = proy2;
                    int j = oUserTable.Add();

                    if (j != 0)
                    {
                        oApp.SetStatusBarMessage("Error" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                    }
                    else
                    {
                        oApp.SetStatusBarMessage("Exito en la inserción", SAPbouiCOM.BoMessageTime.bmt_Medium, false);

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                       

                        oForm.DataSources.DataTables.Item(0).ExecuteQuery("select t0.U_CodigoEDC AS 'Código' ,t0.U_NombreEDC as 'Nombre',t0.U_ProyectoEDC as 'Proyecto',t1.U_NombrePyto as 'Descripción' from [@EDICIONESCOL] as t0   left join [@PROYECTOSCOSTE] as t1 on t0.U_ProyectoEDC = (CAST (t1.U_CentroPyto AS VARCHAR) + '' + CAST (t1.U_DeptoPyto AS VARCHAR) + '' + CAST (t1.U_CodigoPyto AS VARCHAR)) ");
                        Grid2.DataTable = oForm.DataSources.DataTables.Item("dted");
                        Grid2.DataTable.Rows.Add(1);
                        for (int i = 1; i <= this.Grid2.DataTable.Rows.Count; i += 1)
                        {

                            if (i < this.Grid2.DataTable.Rows.Count)
                            {

                                Grid2.Rows.SelectedRows.Add(i);
                            }

                        }
                        gridreco();
                        gridset();
                        RowNumberGrid(Grid2);
                        BubbleEvent = false;



                    }
                }

            }
            //}




        }
        //evento para que cuando se presione una celda se pueda selecciona todo la fila
        private void Grid2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid2.Rows.SelectedRows.Clear();

            Grid2.Rows.SelectedRows.Add(pVal.Row);

            
           
            //Grid2.DataTable.GetValue("CODE", pVal.Row);
            //Grid2.DataTable = oForm.DataSources.DataTables.Item("dted");
            // string sdo = (string)Grid2.DataTable.GetValue("Nombre", Grid2.GetDataTableRowIndex(pVal.Row));

        }

        private void Grid2_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if ((pVal.ColUID == "Proyecto"))
            {

                int dtRow = Grid2.GetDataTableRowIndex(pVal.Row);
                SAPbouiCOM.ComboBoxColumn oCBC = (SAPbouiCOM.ComboBoxColumn)Grid2.Columns.Item("Proyecto");

                string ValSelec = oCBC.GetSelectedValue(pVal.Row).Value; // Selecciona El Valor (U_CodigoSS)
                string DesSelec = oCBC.GetSelectedValue(pVal.Row).Description; // Selecciona la Descripcion (U_NombreSS)
                //mostrar la descripción en el combobox
                //oCBC.DisplayType = (SAPbouiCOM.BoComboDisplayType.cdt_Description);
                // Para Asignar el Valor a una celda del grid se puede asi:
                Grid2.DataTable.SetValue("Descripción", dtRow, DesSelec);


                //O tambien asi
                //Grid0.DataTable.Columns.Item("CODSUPRA").Cells(pVal.Row).value = ValSelec;

            }

        }

  

      

       
  

        //private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{

          
        
                 
                        
        //}

        }
                  
                 
         


              

                    



                  


             

        







     
    
}
