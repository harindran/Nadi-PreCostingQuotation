using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using PreCosting_Quotation.Common;

namespace PreCosting_Quotation.Masters
{
    [FormAttribute("PCQTMSTR", "Masters/FrmPreCostQuoteMaster.b1f")]
    class FrmPreCostQuoteMaster : UserFormBase
    {
        public static SAPbouiCOM.Form objform;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsComponent, odbdsBoughtOut, odbdsTestProp, odbdsOverheadProp, odbdsAttachment,odbdsSpare,odbdsContingency,odbdsWarranty,odbdsNegotiation;
        private string strSQL, strQuery;
        private SAPbobsCOM.Recordset objRs;
        SAPbouiCOM.ISBOChooseFromListEventArg pCFL;

        public FrmPreCostQuoteMaster(string Code)
        {
            try
            {
                if (Code == "") return;
                objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                EditText0.Value = Code;
                objform.Items.Item("1").Click();
                //objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
       
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lcode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tcode").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lfanmod").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tfanmod").Specific));
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.EditText1.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText1_ChooseFromListBefore);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lclass").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tclass").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("larr").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tarr").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lmotkw").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("tmotkw").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lpole").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("tpole").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chkactive").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fcomp").Specific));
            this.Folder0.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder0_PressedAfter);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("fbout").Specific));
            this.Folder1.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder1_PressedAfter);
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("ftprop").Specific));
            this.Folder2.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder2_PressedAfter);
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("foprop").Specific));
            this.Folder3.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder3_PressedAfter);
            this.Folder4 = ((SAPbouiCOM.Folder)(this.GetItem("fattach").Specific));
            this.Folder4.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder4_PressedAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mcomp").Specific));
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("mbout").Specific));
            this.Matrix1.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix1_LinkPressedBefore);
            this.Matrix1.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix1_ValidateAfter);
            this.Matrix1.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix1_ChooseFromListAfter);
            this.Matrix1.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix1_ChooseFromListBefore);
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("mtprop").Specific));
            this.Matrix2.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix2_LinkPressedBefore);
            this.Matrix2.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix2_ValidateAfter);
            this.Matrix2.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix2_ChooseFromListAfter);
            this.Matrix2.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix2_ChooseFromListBefore);
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("moprop").Specific));
            this.Matrix3.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix3_LinkPressedBefore);
            this.Matrix3.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix3_ValidateAfter);
            this.Matrix3.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix3_ChooseFromListAfter);
            this.Matrix3.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix3_ChooseFromListBefore);
            this.Matrix4 = ((SAPbouiCOM.Matrix)(this.GetItem("mattach").Specific));
            this.Matrix4.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.Matrix4_DoubleClickAfter);
            this.Matrix4.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix4_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btnbrowse").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("btndisp").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("btndel").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("tentry").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkfanmod").Specific));
            this.LinkedButton0.PressedBefore += new SAPbouiCOM._ILinkedButtonEvents_PressedBeforeEventHandler(this.LinkedButton0_PressedBefore);
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("lname").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("tname").Specific));
            this.Folder5 = ((SAPbouiCOM.Folder)(this.GetItem("fspare").Specific));
            this.Folder5.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder5_PressedAfter);
            this.Matrix5 = ((SAPbouiCOM.Matrix)(this.GetItem("mspare").Specific));
            this.Matrix5.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix5_ValidateAfter);
            this.Matrix5.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix5_LinkPressedBefore);
            this.Matrix5.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix5_ChooseFromListAfter);
            this.Matrix5.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix5_ChooseFromListBefore);
            this.Folder6 = ((SAPbouiCOM.Folder)(this.GetItem("fccharge").Specific));
            this.Folder6.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder6_PressedAfter);
            this.Folder7 = ((SAPbouiCOM.Folder)(this.GetItem("fwcharge").Specific));
            this.Folder7.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder7_PressedAfter);
            this.Folder8 = ((SAPbouiCOM.Folder)(this.GetItem("fncharge").Specific));
            this.Folder8.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder8_PressedAfter);
            this.Matrix6 = ((SAPbouiCOM.Matrix)(this.GetItem("mccharge").Specific));
            this.Matrix6.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix6_ValidateAfter);
            this.Matrix6.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix6_LinkPressedBefore);
            this.Matrix6.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix6_ChooseFromListAfter);
            this.Matrix6.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix6_ChooseFromListBefore);
            this.Matrix7 = ((SAPbouiCOM.Matrix)(this.GetItem("mwcharge").Specific));
            this.Matrix7.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix7_ValidateAfter);
            this.Matrix7.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix7_LinkPressedBefore);
            this.Matrix7.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix7_ChooseFromListAfter);
            this.Matrix7.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix7_ChooseFromListBefore);
            this.Matrix8 = ((SAPbouiCOM.Matrix)(this.GetItem("mncharge").Specific));
            this.Matrix8.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix8_ValidateAfter);
            this.Matrix8.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix8_LinkPressedBefore);
            this.Matrix8.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix8_ChooseFromListAfter);
            this.Matrix8.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix8_ChooseFromListBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private void OnCustomInitialize()
        {
            try
            {
                CheckBox0.Checked = false;
                objform.Left = (clsModule.objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2;//form.Left + 50;
                objform.Top = (clsModule.objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4;// form.Top + 50;     
                objform.ClientHeight = Button0.Item.Top + Button0.Item.Height + 10;
                objform.ClientWidth = objform.Width;
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR");
                odbdsComponent = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR1");//Components 
                odbdsBoughtOut = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR2"); //BoughtOut
                odbdsTestProp = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR3"); //Test Prop
                odbdsOverheadProp = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR4"); //Overhead Prop
                odbdsAttachment = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR5"); //Attachments             
                odbdsSpare = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR6"); //Spare             
                odbdsContingency = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR7"); //Contingency              
                odbdsWarranty = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR8"); //Warranty             
                odbdsNegotiation = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR9"); //Negotiation             
                Matrix0.Item.Enabled = false; Matrix1.Item.Enabled = false; Matrix2.Item.Enabled = false; Matrix3.Item.Enabled = false; Matrix4.Item.Enabled = false; Matrix5.Item.Enabled = false;
                Matrix6.Item.Enabled = false; Matrix7.Item.Enabled = false; Matrix8.Item.Enabled = false; 
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btnbrowse", true, false, true);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btndisp", true, false, true);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "btndel", true, false, true);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tcode", false, true, false);

                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tfanmod", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tclass", false, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tarr", false, true, false);
                Folder4.Item.Left = Folder8.Item.Left + Folder8.Item.Width+ 1;  //Attachments Tab Position
                Matrix4.Columns.Item("srcpath").Visible = false;
                Matrix4.Columns.Item("fileext").Visible = false;
                Folder0.Item.Click();
                objform.EnableMenu("779", true);
                
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
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Folder Folder4;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.Matrix Matrix4;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;     
        private SAPbouiCOM.Button Button4;     
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.Folder Folder5;
        private SAPbouiCOM.Matrix Matrix5;
        private SAPbouiCOM.Folder Folder6;
        private SAPbouiCOM.Folder Folder7;
        private SAPbouiCOM.Folder Folder8;
        private SAPbouiCOM.Matrix Matrix6;
        private SAPbouiCOM.Matrix Matrix7;
        private SAPbouiCOM.Matrix Matrix8;

        #endregion

        #region Header & Form Events

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("PCQTMSTR", pVal.FormTypeCount);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void EditText1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                ChooseFromList_Condition("cflfanmod", "U_Active", "Y");               

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void EditText1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
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
                        odbdsHeader.SetValue("U_ModelCode", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                        odbdsHeader.SetValue("U_Class", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_Class").Cells.Item(0).Value));
                        odbdsHeader.SetValue("U_Arrangement", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_Arrangement").Cells.Item(0).Value));
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

        private void LinkedButton0_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                FrmFanModel model = new FrmFanModel(EditText1.Value);
                model.Show();
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
                if (EditText1.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Fan Model Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText1.Item.Click();
                    return;
                }
                if (EditText2.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Class is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; 
                    return;
                }
                if (EditText3.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Arrangement is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; 
                    return;
                }
                strSQL = string.Concat(EditText1.Value, "_", EditText2.Value, "_", EditText3.Value);
                EditText0.Value = strSQL;
                
                if (Matrix0.VisualRowCount == 0 || ((SAPbouiCOM.EditText)Matrix0.Columns.Item("compcode").Cells.Item(1).Specific).String=="")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Component Row Data is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false;
                    return;
                }
                if (EditText9.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Name is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false;
                    return;
                }
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix0, "compcode");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix1, "bocode");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix2, "testproc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix3, "opropc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix5, "sparc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix6, "contc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix7, "wcode");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix8, "ngcode");
                
            }
            catch (Exception)
            {
            }

        }

        #endregion

        #region Components

        private void Folder0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mcomp";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "compcode", "#");
                Matrix0.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Components

        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "compcode":
                        ChooseFromList_Condition("cflcomp", "U_Active", "Y");
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
                    case "compcode":
                        Matrix0.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsComponent.SetValue("U_CompCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsComponent.SetValue("U_CompName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
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

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "compcode":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "compcode", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix0.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmComponent component = new FrmComponent(((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                component.Show();
            }
            catch (Exception)
            {
            }
        }

        #endregion

        #region BoughtOut

        private void Folder1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mcomp";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "bocode", "#");
                Matrix1.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //BoughtOut

        private void Matrix1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "compcode":
                        ChooseFromList_Condition("cflbout", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "bocode":
                        Matrix1.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsBoughtOut.SetValue("U_BOutCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsBoughtOut.SetValue("U_BOutName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix1.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }

        }

        private void Matrix1_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "bocode":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "bocode", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix1.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix1_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmBoughtOut component = new FrmBoughtOut(((SAPbouiCOM.EditText)Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                component.Show();
            }
            catch (Exception)
            {
            }
        }

        #endregion

        #region Test Property

        private void Folder2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mtprop";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "testproc", "#");
                Matrix2.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        } //Test Property

        private void Matrix2_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "testproc":
                        ChooseFromList_Condition("cfltprop", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix2_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "testproc":
                        Matrix2.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsTestProp.SetValue("U_TPropCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsTestProp.SetValue("U_TPropName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix2.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }

        }

        private void Matrix2_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "testproc":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "testproc", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix2.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix2_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmTestProperty testProperty = new FrmTestProperty(((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                testProperty.Show();
            }
            catch (Exception)
            {
            }

        }

        #endregion

        #region Packing Charges

        private void Folder3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "moprop";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "opropc", "#");
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        } //Overhead Property

        private void Matrix3_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "opropc":
                        ChooseFromList_Condition("cfloprop", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Matrix3_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "opropc":
                        Matrix3.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsOverheadProp.SetValue("U_OPropCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsOverheadProp.SetValue("U_OPropName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix3.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }

        }

        private void Matrix3_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "opropc":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "opropc", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix3_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmPackingCharges overheadProperty = new FrmPackingCharges(((SAPbouiCOM.EditText)Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                overheadProperty.Show();
            }
            catch (Exception)
            {
            }

        }

        #endregion

        #region Spares 

        private void Folder5_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mspare";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix5, "sparc", "#");
                Matrix5.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Spare Master

        private void Matrix5_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "sparc":
                        ChooseFromList_Condition("cflspare", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        } //Spare Master

        private void Matrix5_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "sparc":
                        Matrix5.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsSpare.SetValue("U_SparCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsSpare.SetValue("U_SparName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix5.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix5.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }

        }

        private void Matrix5_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix5.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmSpareMaster spareMaster = new FrmSpareMaster(((SAPbouiCOM.EditText)Matrix5.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                spareMaster.Show();
            }
            catch (Exception)
            {
            }

        }

        private void Matrix5_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "sparc":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix5, "sparc", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix5.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }
        }

        #endregion

        #region Contingency 

        private void Folder6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mccharge";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix6, "contc", "#");
                Matrix6.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Contingency 

        private void Matrix6_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "contc":
                        ChooseFromList_Condition("ccchrge", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        } //Contingency Charges

        private void Matrix6_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "contc":
                        Matrix6.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsContingency.SetValue("U_ContCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsContingency.SetValue("U_ContName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix6.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }
        } //Contingency Charges

        private void Matrix6_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmContingencyCharges contingencyCharges = new FrmContingencyCharges(((SAPbouiCOM.EditText)Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                contingencyCharges.Show();
            }
            catch (Exception)
            {
            }
        }

        private void Matrix6_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "contc":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix6, "contc", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix6.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        }

        #endregion

        #region Warranty 

        private void Folder7_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mwcharge";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix7, "wcode", "#");
                Matrix7.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Warranty 

        private void Matrix7_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "wcode":
                        ChooseFromList_Condition("cwchrge", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        } //Warranty Charges

        private void Matrix7_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "wcode":
                        Matrix7.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsWarranty.SetValue("U_WarrCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsWarranty.SetValue("U_WarrName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix7.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }

        } //Warranty Charges

        private void Matrix7_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmWarrantyCharges spareMaster = new FrmWarrantyCharges(((SAPbouiCOM.EditText)Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                spareMaster.Show();
            }
            catch (Exception)
            {
            }
        }

        private void Matrix7_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "wcode":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix7, "wcode", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix7.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }
        }

        #endregion

        #region Negotiation 

        private void Folder8_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mncharge";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && objform.Mode != SAPbouiCOM.BoFormMode.fm_VIEW_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix8, "ngcode", "#");
                Matrix8.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Negotiation 

        private void Matrix8_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true) return;
                switch (pVal.ColUID)
                {
                    case "ngcode":
                        ChooseFromList_Condition("cflspare", "U_Active", "Y");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        } //Negotiation Charges

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                
            }
            catch (Exception ex)
            {

            }

        }

        private void Matrix8_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                switch (pVal.ColUID)
                {
                    case "ngcode":
                        Matrix8.FlushToDataSource();
                        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                        if (pCFL.SelectedObjects != null)
                        {
                            try
                            {
                                odbdsNegotiation.SetValue("U_NegoCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value));
                                odbdsNegotiation.SetValue("U_NegoName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value));
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        Matrix8.LoadFromDataSource();
                        break;
                    default:
                        break;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

            }
            catch (Exception)
            {
            }
        } //Negotiation Charges

        private void Matrix8_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                FrmNegotiationCharges spareMaster = new FrmNegotiationCharges(((SAPbouiCOM.EditText)Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                spareMaster.Show();
            }
            catch (Exception)
            {
            }
        }

        private void Matrix8_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "ngcode":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix8, "ngcode", "#");
                        break;
                }
                objform.Freeze(true);
                Matrix8.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }
        }

        #endregion

        #region Attachments       

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.SetAttachMentFile(objform, odbdsHeader, Matrix4, odbdsAttachment);
                if (Matrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) == -1)
                {
                    objform.Items.Item("btndisp").Enabled = false;
                    objform.Items.Item("btndel").Enabled = false;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Browse Attachment 
             
        private void Button3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix4, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }  //Display Attachment

        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.DeleteRowAttachment(objform, Matrix4, odbdsAttachment, Matrix4.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        } //Delete Attachment        

        private void Folder4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mattach";
                Matrix4.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }  //Attachment Tab

        private void Matrix4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Matrix4.SelectRow(pVal.Row, true, false);
                if (Matrix4.IsRowSelected(pVal.Row) == true)
                {
                    objform.Items.Item("btndisp").Enabled = true;
                    objform.Items.Item("btndel").Enabled = true;
                }
            }
            catch (Exception ex)
            {

            }

        } //Attachment Matrix

        private void Matrix4_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix4, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {

            }

        } //Attachment Matrix

        #endregion

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

        
    }
}
