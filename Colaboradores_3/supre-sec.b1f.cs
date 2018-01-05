using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.supre_sec", "supre-sec.b1f")]
    class supre_sec : UserFormBase
    {
        public supre_sec()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GR_SUPRA").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
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



        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           

        }

        private void OnCustomInitialize()
        {
            RowNumberGrid(Grid0);
        }

        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
    }
}
