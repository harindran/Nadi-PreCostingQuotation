using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using PreCosting_Quotation.Masters;
using PreCosting_Quotation.Transactions;

namespace PreCosting_Quotation.Common
{
    class clsAddon
    {
        public SAPbouiCOM.Application objapplication;
        public SAPbobsCOM.Company objcompany;
        public clsMenuEvent objmenuevent;
        public clsRightClickEvent objrightclickevent;
        public clsGlobalMethods objglobalmethods;
        public clsGlobalVariables objGlobalVariables;
        private SAPbouiCOM.Form objform;
        string strsql= "";
        private SAPbobsCOM.Recordset objrs;
        bool print_close = false;
        public bool HANA = true;
        //public bool HANA = false;
        private SAPbouiCOM.Form tempform = null;
        int CurRow = -1;
        public string[] HWKEY   =  { "L1653539483" };

        #region Constructor
        public clsAddon()
        {
            
        }
        #endregion

        public void Intialize(string[] args)
        {
            try
            {
                Application oapplication;
                if ((args.Length < 1))
                    oapplication = new Application();
                else
                    oapplication = new Application(args[0]);
                objapplication = Application.SBO_Application;
                if (isValidLicense())
                {
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objcompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                    Create_DatabaseFields(); // UDF & UDO Creation Part    
                    Menu(); // Menu Creation Part
                    Create_Objects(); // Object Creation Part
                    Add_Authorizations(); //User Permissions
                    clsModule.objaddon.objglobalmethods.addReport_Layouttype("PreCost Layout", "PreCost Quotation");
                    objapplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(objapplication_AppEvent);
                    objapplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(objapplication_MenuEvent);
                    objapplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objapplication_ItemEvent);
                    objapplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(ref FormDataEvent);
                    objapplication.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(objapplication_LayoutKeyEvent);
                    //objapplication.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(objapplication_ProgressBarEvent);
                    //objapplication.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(objapplication_StatusBarEvent);
                    objapplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(objapplication_RightClickEvent);

                    objapplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    oapplication.Run();
                }
                else
                {
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                    //throw new Exception(objcompany.GetLastErrorDescription());
                }
            }
            // System.Windows.Forms.Application.Run()
            catch (Exception ex)
            {
                objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }      
        
        public bool isValidLicense()
        {
            try
            {
                if (HANA)
                {
                    try
                    {
                        if (objapplication.Forms.ActiveForm.TypeCount > 0)
                        {
                            for (int i = 0; i <= objapplication.Forms.ActiveForm.TypeCount - 1; i++)
                                objapplication.Forms.ActiveForm.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                // If Not HANA Then
                // objapplication.Menus.Item("1030").Activate()
                // End If
                objapplication.Menus.Item("257").Activate();
                SAPbouiCOM.EditText objedit= (SAPbouiCOM.EditText)objapplication.Forms.ActiveForm.Items.Item("79").Specific;
              
                string CrrHWKEY = objedit.Value.ToString();
                objapplication.Forms.ActiveForm.Close();

                for (int i = 0; i <= HWKEY.Length - 1; i++)
                {
                    //string key = HWKEY[i];
                    if (HWKEY[i] == CrrHWKEY)
                    {
                        return true;
                    }
                        
                }
                
                System.Windows.Forms.MessageBox.Show("Installing Add-On failed due to License mismatch");
                //objapplication.MessageBox("Installing Add-On failed due to License mismatch", 1, "Ok", "", "");
                //Interaction.MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management");

                return false;
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return true;
        }

        public void Create_Objects()
        {
            objmenuevent = new clsMenuEvent();
            objrightclickevent = new clsRightClickEvent();
            objglobalmethods = new clsGlobalMethods();
            objGlobalVariables = new clsGlobalVariables();
        }

        private void Create_DatabaseFields()
        {
            // If objapplication.Company.UserName.ToString.ToUpper <> "MANAGER" Then

            // If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            var objtable = new clsTable();
            objtable.FieldCreation();
            // End If

        }

        public void Add_Authorizations()
        {
            try
            {
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Altrocks Tech", "ATPL_ADD-ON", "", "", 'Y');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Pre Costing Quote", "ATPL_PRECOST", "", "ATPL_ADD-ON", 'Y');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Master", "ATPL_PCMSTR", "", "ATPL_PRECOST", 'Y');

                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Fan Model", "ATPL_PCMODEL", "PCMODEL", "ATPL_PCMSTR", 'Y');//Fan Model
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Component", "ATPL_PCCOMP", "PCCOMP", "ATPL_PCMSTR", 'Y');//Component
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("BoughtOut", "ATPL_PCBOUT", "PCBOUT", "ATPL_PCMSTR", 'Y');//Boughtout
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Test Property", "ATPL_PCTPROP", "PCTPROP", "ATPL_PCMSTR", 'Y');//Test Property
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Packing Charges", "ATPL_PCOPROP", "PCOPROP", "ATPL_PCMSTR", 'Y');//Packing Charges
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Spares", "ATPL_PCSPAR", "PCSPAR", "ATPL_PCMSTR", 'Y');//Spares
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Contingency Charges", "ATPL_PCCCGE", "PCCCGE", "ATPL_PCMSTR", 'Y');//Contingency Charges
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Warranty Charges", "ATPL_PCWCGE", "PCWCGE", "ATPL_PCMSTR", 'Y');//Warranty Charges
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Negotiation Charges", "ATPL_PCNCGE", "PCNCGE", "ATPL_PCMSTR", 'Y');//Negotiation Charges                
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("PreCost Quote Master", "ATPL_PCQTMSTR", "PCQTMSTR", "ATPL_PCMSTR", 'Y');//Pre-Cost Quote Master

                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Transaction", "ATPL_PCTRAN", "", "ATPL_PRECOST", 'Y');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("PreCost Quotation", "ATPL_PCQUOTE", "PCQUOTE", "ATPL_PCTRAN", 'Y');//Pre-Cost Quote Transaction

            }
            catch (Exception ex)
            {
            }
        }

        #region Menu Creation Details

        private void Menu()
        {
            int Menucount = 1;
            if (objapplication.Menus.Item("43520").SubMenus.Exists("PRECOST") )
                return;
            Menucount = 5;// objapplication.Menus.Item("8448").SubMenus.Count;
            CreateMenu("", Menucount, "Pre Costing Quote", SAPbouiCOM.BoMenuType.mt_POPUP, "PRECOST", "43520"); Menucount = 1;

            //CreateMenu("", Menucount, "SetUp", SAPbouiCOM.BoMenuType.mt_POPUP, "PCSETUP", "PRECOST"); Menucount += 1;
            CreateMenu("", Menucount, "Master", SAPbouiCOM.BoMenuType.mt_POPUP, "PCMSTR", "PRECOST"); Menucount = 1;//Menu Inside  
            CreateMenu("", Menucount, "Transaction", SAPbouiCOM.BoMenuType.mt_POPUP, "PCTRAN", "PRECOST"); Menucount = 1;//Menu Inside  

            CreateMenu("", Menucount, "Fan Model", SAPbouiCOM.BoMenuType.mt_STRING, "PCMODEL", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Components", SAPbouiCOM.BoMenuType.mt_STRING, "PCCOMP", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Bought Out", SAPbouiCOM.BoMenuType.mt_STRING, "PCBOUT", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Test Properties", SAPbouiCOM.BoMenuType.mt_STRING, "PCTPROP", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Packing Charges", SAPbouiCOM.BoMenuType.mt_STRING, "PCOPROP", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Spares Master", SAPbouiCOM.BoMenuType.mt_STRING, "PCSPAR", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Contingency Charges", SAPbouiCOM.BoMenuType.mt_STRING, "PCCCGE", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Warranty Charges", SAPbouiCOM.BoMenuType.mt_STRING, "PCWCGE", "PCMSTR"); Menucount += 1;
            CreateMenu("", Menucount, "Negotiation Charges", SAPbouiCOM.BoMenuType.mt_STRING, "PCNCGE", "PCMSTR"); Menucount += 1;            
            CreateMenu("", Menucount, "Pre Cost Quote Master", SAPbouiCOM.BoMenuType.mt_STRING, "PCQTMSTR", "PCMSTR"); Menucount += 1;
                        
            CreateMenu("", Menucount, "Pre Cost Quotation", SAPbouiCOM.BoMenuType.mt_STRING, "PCQUOTE", "PCTRAN"); //Menucount += 1;
           
        }

        private void CreateMenu(string ImagePath, int Position, string DisplayName, SAPbouiCOM.BoMenuType MenuType, string UniqueID, string ParentMenuID)
        {
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuPackage;
                SAPbouiCOM.MenuItem parentmenu;
                parentmenu = objapplication.Menus.Item(ParentMenuID);
                if (parentmenu.SubMenus.Exists(UniqueID.ToString()))
                    return;
                oMenuPackage =(SAPbouiCOM.MenuCreationParams) objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuPackage.Image = ImagePath;
                oMenuPackage.Position = Position;
                oMenuPackage.Type = MenuType;
                oMenuPackage.UniqueID = UniqueID;
                oMenuPackage.String = DisplayName;
                parentmenu.SubMenus.AddEx(oMenuPackage);
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            // Return ParentMenu.SubMenus.Item(UniqueID)
        }

        #endregion

        public bool FormExist(string FormID)
        {
            bool FormExistRet = false;
            try
            {
                FormExistRet = false;
                foreach (SAPbouiCOM.Form uid in clsModule.objaddon.objapplication.Forms)
                {
                    if (uid.TypeEx == FormID)
                    {
                        FormExistRet = true;
                        break;
                    }
                }
                if (FormExistRet)
                {
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Visible = true;
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Select();
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return FormExistRet;

        }

        #region VIRTUAL FUNCTIONS
        //public virtual void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        //{ }

        //public virtual void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{ }

        //public virtual void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool BubbleEvent)
        //{ }        

        //public virtual void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{ }

      
        #endregion

        #region ItemEvent

        private void objapplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)  
        {
            BubbleEvent = true;
            try
            {                
                switch (pVal.FormTypeEx)
                {
                    case "":
                        //objInvoice.Item_Event(FormUID, ref pVal,ref BubbleEvent);
                        break;
                    
                }
                if (pVal.BeforeAction)
                {
                    {
                        objform = objapplication.Forms.Item(FormUID);
                        
                        switch (pVal.EventType)
                            {
                            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                                
                                break;

                            case SAPbouiCOM.BoEventTypes.et_CLICK:
                                if (pVal.FormTypeEx == "PCQUOTE" && clsModule.objaddon.objGlobalVariables.bModal == true)
                                {
                                    clsModule.objaddon.objapplication.Forms.Item("PCQTETRAN").Select(); BubbleEvent = false;
                                }
                                break;

                            case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                                if (FormUID == "PCQTETRAN")
                                {
                                    clsModule.objaddon.objGlobalVariables.bModal = false;
                                }
                                break;
                            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                                //if ((pVal.FormTypeEx == "-392"|| pVal.FormTypeEx == "-393") && (pVal.ItemUID== "U_RevRecDE" || pVal.ItemUID == "U_RevRecDN") && pVal.FormMode!=0 ) //0-Find Mode
                                //{
                                //    BubbleEvent = false;
                                //}
                                    break;
                                                        
                        }
                    }

                }
                else
                {
                    switch (pVal.EventType)
                    {                       
                       
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                           
                            break;
                    }
                }
                
            }
            catch (Exception ex) {
                //objapplication.StatusBar.SetText("objapplication_ItemEvent" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
           
           
        }

        #endregion

        #region FormDataEvent

        private void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (BusinessObjectInfo.FormTypeEx)
                {           
                    case ""://ClsARInvoice.Formtype:                    
                        //objInvoice.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);
                    break;
                }

            }
            catch (Exception)
            {

            }
            

        }

        #endregion
        
        #region MenuEvent

        private void objapplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;
            switch (pVal.MenuUID)
            {
                case "1281":
                case "1282":
                case "1283":
                case "1284":
                case "1285":
                case "1286":
                case "1287":
                case "1300":
                case "1288":
                case "1289":
                case "1290":
                case "1291":
                case "1304":
                case "1292":
                case "1293":
                    objmenuevent.MenuEvent_For_StandardMenu(ref pVal, ref BubbleEvent);
                    break;
                case "PCMODEL": //Fan Model
                case "PCCOMP": //Component
                case "PCBOUT": //Boughtout
                case "PCTPROP": //Test Property
                case "PCOPROP": //Overhead Property
                case "PCQTMSTR": //Pre-Cost Quote Master
                case "PCQUOTE": //Pre-Cost Quote Selection
                case "PCQTETRAN": //Pre-Cost Quote Transaction
                case "PCSPAR": //Spare Master
                case "PCCCGE": //Contingency Charges Master
                case "PCNCGE": //Negotiation Charges Master
                case "PCWCGE": //Warranty Charges Master
                    MenuEvent_For_FormOpening(ref pVal, ref BubbleEvent);
                    break;

            }


        }

        public void MenuEvent_For_FormOpening(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {                
                if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "PCMODEL": //Fan Model
                            { 
                            FrmFanModel activeform = new FrmFanModel("");
                            activeform.Show();
                            }
                            break;
                        case "PCCOMP": //Component
                            {
                                FrmComponent activeform = new FrmComponent("");
                                activeform.Show();
                            }
                            break;
                        case "PCBOUT": //Boughtout
                            {
                                FrmBoughtOut activeform = new FrmBoughtOut("");
                                activeform.Show();
                            }
                            break;
                        case "PCTPROP": //Test Property
                            {
                                FrmTestProperty activeform = new FrmTestProperty("");
                                activeform.Show();
                            }
                            break;
                        case "PCOPROP": //Packing Charges
                            {
                                FrmPackingCharges activeform = new FrmPackingCharges("");
                                activeform.Show();
                            }
                            break;
                        case "PCQTMSTR": //Pre-Cost Quote Master
                            {
                                FrmPreCostQuoteMaster activeform = new FrmPreCostQuoteMaster("");
                                activeform.Show();
                            }
                            break;
                        case "PCQUOTE": //Pre-Cost Quote Transaction
                            {
                                FrmPCQuoteSelection activeform = new FrmPCQuoteSelection();                                
                                activeform.Show();
                            }                            
                            break;
                        case "PCSPAR": //Spare Master
                            {
                                FrmSpareMaster activeform = new FrmSpareMaster("");
                                activeform.Show();
                            }
                            break;
                        case "PCCCGE": //Contingency Charges Master
                            {
                                FrmContingencyCharges activeform = new FrmContingencyCharges("");
                                activeform.Show();
                            }
                            break;
                        case "PCNCGE": //Negotiation Charges Master
                            {
                                FrmNegotiationCharges activeform = new FrmNegotiationCharges("");
                                activeform.Show();
                            }
                            break;
                        case "PCWCGE": //Warranty Charges Master
                            {
                                FrmWarrantyCharges activeform = new FrmWarrantyCharges("");
                                activeform.Show();
                            }
                            break;

                    }

                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
        }

        #endregion

        #region RightClickEvent

        private void objapplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "PCMODEL": //Fan Model
                    case "PCCOMP": //Component
                    case "PCBOUT": //Boughtout
                    case "PCTPROP": //Test Property
                    case "PCOPROP": //Overhead Property
                    case "PCQTMSTR": //Pre-Cost Quote Master
                    case "PCQUOTE": //Pre-Cost Quote Selection
                    case "PCQTETRAN": //Pre-Cost Quote Transaction
                        objrightclickevent.RightClickEvent(ref eventInfo, ref BubbleEvent);
                    break;
                }

            }
            catch (Exception ex) { }
            
        }

        #endregion

        #region LayoutEvent
        
        public void objapplication_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion

        #region AppEvent

        private void objapplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {          
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    //objapplication.MessageBox("A Shut Down Event has been caught" + Environment.NewLine + "Terminating Add On...", 1, "Ok", "", "");
                    try
                    {
                        Remove_Menu(new[] { "43520,PRECOST" });
                        DisConnect_Addon();
                        //System.Windows.Forms.Application.Exit();
                        //if (objapplication != null)
                        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                        //if (objcompany != null)
                        //{
                        //    if (objcompany.Connected)
                        //        objcompany.Disconnect();
                        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                        //}                        
                        //GC.Collect();                        
                        ////Environment.Exit(0);
                    }
                    catch (Exception ex)
                    {
                    }               
                    break;
               
            }
        }

        private void DisConnect_Addon()
        {
            try
            {
                if (clsModule.objaddon.objapplication.Forms.Count > 0)
                {
                    try
                    {
                        for (int frm = clsModule.objaddon.objapplication.Forms.Count - 1; frm >= 0; frm--)
                        {
                            if (clsModule.objaddon.objapplication.Forms.Item(frm).IsSystem == true) continue;
                            clsModule.objaddon.objapplication.Forms.Item(frm).Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    } 
                }
                if (objcompany.Connected)
                    objcompany.Disconnect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                objcompany = null;
                GC.Collect();
                System.Windows.Forms.Application.Exit();
                System.Environment.Exit(0);
                //Environment.Exit(0);
            }
            catch (Exception ex)
            {

            }
        }

        private void Remove_Menu(string[] MenuID = null)
        {
            try
            {
                string[] split_char;
                if (MenuID != null)
                {
                    if (MenuID.Length > 0)
                    {
                        for (int i = 0, loopTo = MenuID.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(MenuID[i]))
                                continue;
                            split_char = MenuID[i].Split(Convert.ToChar(","));
                            if (split_char.Length != 2)
                                continue;
                            if (clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.Exists(split_char[1]))
                                clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.RemoveEx(split_char[1]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }         

        }
        
        #endregion


        }


}
