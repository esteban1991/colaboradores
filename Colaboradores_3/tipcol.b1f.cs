using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    [FormAttribute("Colaboradores_3.tipcol", "tipcol.b1f")]
    class tipcol : UserFormBase
    {
        public tipcol()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GR_TICOL").Specific));
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

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button1;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
          

        }

        private SAPbouiCOM.Grid Grid0;
    }
}
