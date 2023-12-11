using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using PreCosting_Quotation.Common;

namespace PreCosting_Quotation.Masters
{
    [FormAttribute("PCOPROP", "Masters/FrmPackingCharges.b1f")]
    class FrmPackingCharges : UserFormBase
    {
        public FrmPackingCharges(string Code)
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
        public static SAPbouiCOM.Form objform;
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
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lname").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tname").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chkactive").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tentry").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {
            try
            {
                CheckBox0.Checked = false;
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tcode", true, true, false);
                objform.Left = (clsModule.objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2;//form.Left + 50;
                objform.Top = (clsModule.objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4;// form.Top + 50;
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
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;

        #endregion

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("PCOPROP", pVal.FormTypeCount);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                    clsModule.objaddon.objapplication.StatusBar.SetText("Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText0.Item.Click();
                    return;
                }
                if (EditText1.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Name is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText1.Item.Click();
                    return;
                }
            }
            catch (Exception ex)
            {
            }
           

        }
    }
}
