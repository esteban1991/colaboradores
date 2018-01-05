using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace Colaboradores_3
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "Colaboradores_3";
            oCreationPackage.String = "Colaboradores";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception )
            {

            }




            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Colaboradores_3");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Colaboradores_3.Form1";
                oCreationPackage.String = "Colaboradores";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception )
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Este menu ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }



          


            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Colaboradores_3");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Colaboradores_3.suprasec";
                oCreationPackage.String = "Supra-Secciones";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Este menu ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }


            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Colaboradores_3");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Colaboradores_3.seccionesgr";
                oCreationPackage.String = "Secciones";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Este menu ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Colaboradores_3");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Colaboradores_3.tipcol";
                oCreationPackage.String = "Tipos de Colaboración";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Este menu ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }


            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Colaboradores_3");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Colaboradores_3.edicionesgr";
                oCreationPackage.String = "Ediciones";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Este menu ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }


          
     
           
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Colaboradores_3.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

          
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Colaboradores_3.suprasec")
                {
                    suprasec activeForm = new suprasec();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Colaboradores_3.seccionesgr")
                {
                    seccionesgr activeForm = new seccionesgr();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Colaboradores_3.tipcol")
                {
                    tipcol activeForm = new tipcol();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Colaboradores_3.edicionesgr")
                {
                    edicionesgr activeForm = new edicionesgr();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

         

           
        }
     
      


    }
}
