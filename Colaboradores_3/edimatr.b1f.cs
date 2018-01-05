using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{

    [FormAttribute("Colaboradores_3.edimatr", "edimatr.b1f")]
    class edimatr : UserFormBase
    {
        public SAPbouiCOM.Application oApp;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        //public SAPbobsCOM.UserTable oUserTable;
        //private SAPbouiCOM.Button Button0;
        //private SAPbouiCOM.Button Button1;
        public SAPbouiCOM.Matrix oMatrix; 
        public edimatr()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_2").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
           
            this.OnCustomInitialize();



        }

        private SAPbouiCOM.DBDataSource oDBDataSource; 
        private SAPbouiCOM.UserDataSource oUserDataSource; 
        public void AddDataSourceToForm()
        {

            //****************************************************************
            // every item must be binded to a Data Source
            // prior of binding the data we must add Data sources to the form
            //****************************************************************

            // Add user data sources to the "International Phone" column in the matrix
            oUserDataSource = oForm.DataSources.UserDataSources.Add("CÓDIGO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            // Add DB data sources for the DB bound columns in the matrix
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@EDICIONESCOL");
        } 
        public void Start()
        {
            string dataTableId = "matrix";

            AddDataSourceToForm();
            // add a data table to the form
            var dataTable = oForm.DataSources.DataTables.Add(dataTableId);

            // add a matrix to the form
            Matrix0 = (SAPbouiCOM.Matrix)oForm.Items.Add(dataTableId, SAPbouiCOM.BoFormItemTypes.it_MATRIX).Specific;
            
            Matrix0.Item.Width = 700;
            Matrix0.Item.Height = 300;


            // get some data
            //string sql = @"SELECT ""CardCode"", ""CardName"" FROM ""OCRD"" ";
            string sql = @"SELECT ""U_CodigoEDC"", ""U_NombreEDC"",""U_ProyectoEDC"" FROM ""@EDICIONESCOL""  ";
            //  oApp.SetStatusBarMessage("codigo: " + sql);
            dataTable.ExecuteQuery(sql);

            var columns = Matrix0.Columns;

            // create columns
            columns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            columns.Add("Codigo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            columns.Add("Nombre", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            columns.Add("Proyecto", SAPbouiCOM.BoFormItemTypes.it_EDIT);

            // setup columns
            columns.Item("#").TitleObject.Caption = "#";
            columns.Item("Codigo").TitleObject.Caption = "Código";
            columns.Item("Nombre").TitleObject.Caption = "Nombre";
            columns.Item("Proyecto").TitleObject.Caption = "Proyecto";

            //bind columns to data
            columns.Item("Codigo").DataBind.Bind(dataTableId, "U_CodigoEDC");
            columns.Item("Nombre").DataBind.Bind(dataTableId, "U_NombreEDC");
            columns.Item("Proyecto").DataBind.Bind(dataTableId, "U_ProyectoEDC");

            // load the data into the rows
            Matrix0.LoadFromDataSource();
            Matrix0.AutoResizeColumns();

            // show form on screen
            oForm.Visible = true;
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
            oForm = oApp.Forms.Item("edimatr");
            //Start();
            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("Item_2").Specific));

            oDBDataSource = oForm.DataSources.DBDataSources.Item("@EDICIONESCOL");

            // Ready Matrix to populate data

            oMatrix.Clear();

            oMatrix.AutoResizeColumns();

            // Querying the DB Data source

            oDBDataSource.Query();

            // setting the user data source data

            oMatrix.LoadFromDataSource();

            oMatrix.AutoResizeColumns();

            //int i;
            //for (i = 1; i <= oMatrix.RowCount - 1; i++) ;

            //{

            //    SAPbouiCOM.EditText cellID = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("numer", i);

            //    cellID.String = i.ToString();

            //}
                            

            oForm.Visible = true;

        }


        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Button Button5;

     

    }
}
