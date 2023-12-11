using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using PreCosting_Quotation.Common;
using PreCosting_Quotation.Masters;

namespace PreCosting_Quotation.Transactions
{
    [FormAttribute("PCQUOTE", "Transactions/FrmPCQuoteSelection.b1f")]
    class FrmPCQuoteSelection : UserFormBase
    {
        public FrmPCQuoteSelection()
        {            
        }
        public static SAPbouiCOM.Form objform;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsDetails;
        private string strSQL, strQuery;
        private SAPbobsCOM.Recordset objRs;
        SAPbouiCOM.ISBOChooseFromListEventArg pCFL;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lbpcode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tbpcode").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.EditText0.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText0_ChooseFromListBefore);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lpname").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tbpname").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lindus").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("series").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lno").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tdocnum").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lstat").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbstat").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("ldocdate").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("tdocdate").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fprodmod").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("fldr2").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mprodmod").Specific));
            this.Matrix0.PickerClickedBefore += new SAPbouiCOM._IMatrixEvents_PickerClickedBeforeEventHandler(this.Matrix0_PickerClickedBefore);
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkbpcode").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tentry").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cindus").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("lrefno").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("trefno").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.LayoutKeyBefore += new LayoutKeyBeforeHandler(this.Form_LayoutKeyBefore);

        }


        private void OnCustomInitialize()
        {
            try
            {
                objform.Left = (clsModule.objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2;//form.Left + 50;
                objform.Top = (clsModule.objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4;// form.Top + 50;
                objform.ClientHeight = Button0.Item.Top + Button0.Item.Height + 10;
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_PCQTESEL");
                odbdsDetails = objform.DataSources.DBDataSources.Item("@AT_PCQTESEL1"); //Details 
                clsModule.objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_PCQTESEL");
                ((SAPbouiCOM.EditText)objform.Items.Item("tdocdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "modname", "#");
                strSQL = "Select \"IndCode\",\"IndName\" from OOND";
                clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("cindus").Specific), strSQL, new[] { ","});
                ComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                objform.EnableMenu("1283", false); objform.EnableMenu("1284", false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tbpcode", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tbpname", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cindus", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "series", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocnum", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cmbstat", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocdate", true, true, false);
                Matrix0.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Folder0.Item.Click();

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #region Fields
        
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText EditText2;

        #endregion

        private void EditText0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                ChooseFromList_Condition("cflbpcode", "CardType", "C");

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false)
                    return;
                pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pCFL.SelectedObjects != null)
                {
                    try
                    {
                        odbdsHeader.SetValue("U_BPCode", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value));
                        odbdsHeader.SetValue("U_BPName", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value));
                        strSQL =clsModule.objaddon.objglobalmethods.getSingleValue( "Select \"IndustryC\" from OCRD Where \"CardCode\"='"+ Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value) + "'");
                        if (strSQL == "0") odbdsHeader.SetValue("U_IndustryC", 0, ""); else odbdsHeader.SetValue("U_IndustryC", 0, strSQL);
                    }
                    catch (Exception ex)
                    {
                    }

                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("PCQUOTE", pVal.FormTypeCount);
                clsModule.objaddon.objglobalmethods.setReport(pVal.FormUID, "PreCost Layout");
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("quotetran").Cells.Item(pVal.Row).Specific).String != "") BubbleEvent = false;
                switch (pVal.ColUID)
                {
                    case "modcode":
                        ChooseFromList_Condition("cflpcqte", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "modcode":
                        Matrix0.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsDetails.SetValue("U_ModelCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsDetails.SetValue("U_ModelName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                                odbdsDetails.SetValue("U_Class", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_Class").Cells.Item(0).Value));
                                odbdsDetails.SetValue("U_Arrangement", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_Arrangement").Cells.Item(0).Value));
                                odbdsDetails.SetValue("U_MotorKW", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_MotorKW").Cells.Item(0).Value));
                                odbdsDetails.SetValue("U_Poles", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_Poles").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix0.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }


        }

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "modcode": 
                        FrmPreCostQuoteMaster preCostQuoteMaster = new FrmPreCostQuoteMaster(((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        preCostQuoteMaster.Show();
                        break;
                    case "quotetran":
                        if (!clsModule.objaddon.FormExist("PCQTETRAN"))
                        {
                            FrmPreCostQuoteTransaction preCostQuoteTransaction = new FrmPreCostQuoteTransaction(((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String, objform, null, null, pVal.Row);
                            preCostQuoteTransaction.Show();
                        }                       
                        break;
                        
                    default:
                        break;
                }
                
            }
            catch (Exception)
            {
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                if (EditText0.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Customer Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText0.Item.Click();
                    return;
                }

                if (Matrix0.VisualRowCount == 0 || ((SAPbouiCOM.EditText)Matrix0.Columns.Item("modcode").Cells.Item(1).Specific).String == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Row Data is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false;
                    return;
                }
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix0, "modcode");

            }
            catch (Exception)
            {
            }
        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "modcode":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "modcode", "#");
                        break;
                }
                //objform.Freeze(true);
                //Matrix0.AutoResizeColumns();
                //objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {                
                if (pVal.ActionSuccess==true && objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    clsModule.objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_PCQTESEL");
                    ((SAPbouiCOM.EditText)objform.Items.Item("tdocdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                    ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                    clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "modname", "#");
                    objform.ActiveItem = "tbpcode";
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                int colID = Matrix0.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix0.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix0.VisualRowCount != pVal.Row) Matrix0.SetCellFocus(pVal.Row + 1, colID);
                }
                switch (pVal.ColUID)
                {                    
                    case "quotetran":
                        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("modcode").Cells.Item(pVal.Row).Specific).String == "" || ((SAPbouiCOM.EditText)Matrix0.Columns.Item("quotetran").Cells.Item(pVal.Row).Specific).String != "") return;
                        objform.Freeze(true);
                        FrmPreCostQuoteTransaction preCostQuoteTransaction = new FrmPreCostQuoteTransaction("", objform, odbdsHeader, odbdsDetails, pVal.Row);
                        preCostQuoteTransaction.Show();
                        objform.Freeze(false);
                        break;

                    default:
                        break;
                }                

            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.ColUID)
                {
                    case "quotetran":
                        //if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String != "") BubbleEvent = false;
                        if (pVal.CharPressed == 9 || pVal.CharPressed == 37 || pVal.CharPressed == 38 || pVal.CharPressed == 39 || pVal.CharPressed == 40) return;
                        else BubbleEvent = false;
                        break;
                    case "total":
                        if (pVal.CharPressed == 9 || pVal.CharPressed == 37 || pVal.CharPressed == 38 || pVal.CharPressed == 39 || pVal.CharPressed == 40) return;
                        else BubbleEvent = false;
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix0_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "quotetran" || pVal.ColUID == "total") { objform.EnableMenu("771", false); objform.EnableMenu("773", false); objform.EnableMenu("774", false);  }
                else { objform.EnableMenu("773", true);  }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Matrix0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                Matrix0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix0_PickerClickedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ColUID == "total") { BubbleEvent = false; }
            }
            catch (Exception ex)
            {

            }

        }

        #region Functions

        private void ChooseFromList_Condition(string CFLID, string Alias, string CondVal)
        {
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item(CFLID);
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond;
                var oEmptyConds = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                oCond = oConds.Add();
                oCond.Alias = Alias;// "Postable";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CondVal;// "Y";   

                oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }



        #endregion

        private void Form_LayoutKeyBefore(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                eventInfo.LayoutKey = EditText6.Value;
            }
            catch (Exception ex)
            {
            }

        }

        
    }
}
