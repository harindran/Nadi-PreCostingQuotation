using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using PreCosting_Quotation.Common;
using PreCosting_Quotation.Masters;

namespace PreCosting_Quotation.Transactions
{
    [FormAttribute("PCQTETRAN", "Transactions/FrmPreCostQuoteTransaction.b1f")]
    class FrmPreCostQuoteTransaction : UserFormBase
    {
        public static SAPbouiCOM.Form objform,selectionForm;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsComponent, odbdsBoughtOut, odbdsTestProp, odbdsServiceChrge, odbdsOverheadProp, odbdsAttachment,odbdsSpares, odbdsContingency, odbdsWarranty, odbdsNegotiation;
        private string strSQL;
        private SAPbobsCOM.Recordset objrs;
        //SAPbouiCOM.ISBOChooseFromListEventArg pCFL;
        private int selectionQteLine;

        public FrmPreCostQuoteTransaction(string docEntry, SAPbouiCOM.Form form, SAPbouiCOM.DBDataSource headerdBSource, SAPbouiCOM.DBDataSource linedBSource,int selectionLine)
        {
            try
            {
                selectionQteLine = selectionLine; selectionForm = form;
                if (docEntry == "")
                {                    
                    odbdsHeader.SetValue("U_QteSelNo", 0, headerdBSource.GetValue("DocNum",0));
                    odbdsHeader.SetValue("U_QteSelEnt", 0, headerdBSource.GetValue("DocEntry", 0));
                    odbdsHeader.SetValue("U_BPCode", 0, headerdBSource.GetValue("U_BPCode", 0));
                    odbdsHeader.SetValue("U_BPName", 0, headerdBSource.GetValue("U_BPName", 0));
                    odbdsHeader.SetValue("U_ModelCode", 0, linedBSource.GetValue("U_ModelCode", selectionLine-1));
                    odbdsHeader.SetValue("U_ModelName", 0, linedBSource.GetValue("U_ModelName", selectionLine-1));
                    odbdsHeader.SetValue("U_ModelRow", 0,Convert.ToString(selectionLine));                    
                }
                else
                {
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    EditText6.Item.Enabled = true;
                    EditText6.Value = docEntry;
                    objform.Items.Item("1").Click();
                    //objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Loading details. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    LoadDataByQuery(odbdsHeader.GetValue("U_ModelCode", 0), docEntry);
                    clsModule.objaddon.objapplication.StatusBar.SetText("Details loaded successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                

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
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tbpcode").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lmodcode").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tbpname").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkbpcode").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("series").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lno").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tdocnum").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lstat").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbstat").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("ldocdate").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tdocdate").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fcomp").Specific));
            this.Folder0.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder0_PressedAfter);
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("fbout").Specific));
            this.Folder1.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder1_PressedAfter);
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("ftprop").Specific));
            this.Folder2.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder2_PressedAfter);
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("fschrg").Specific));
            this.Folder3.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder3_PressedAfter);
            this.Folder4 = ((SAPbouiCOM.Folder)(this.GetItem("fochrg").Specific));
            this.Folder4.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder4_PressedAfter);
            this.Folder5 = ((SAPbouiCOM.Folder)(this.GetItem("fattach").Specific));
            this.Folder5.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder5_PressedAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mcomp").Specific));
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lqtesel").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("tqteselno").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("mbout").Specific));
            this.Matrix1.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix1_KeyDownAfter);
            this.Matrix1.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix1_ClickAfter);
            this.Matrix1.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix1_ValidateAfter);
            this.Matrix1.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix1_LinkPressedBefore);
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("mtprop").Specific));
            this.Matrix2.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix2_KeyDownAfter);
            this.Matrix2.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix2_ClickAfter);
            this.Matrix2.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix2_ValidateAfter);
            this.Matrix2.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix2_LinkPressedBefore);
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("mschrg").Specific));
            this.Matrix3.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix3_KeyDownAfter);
            this.Matrix3.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix3_ValidateAfter);
            this.Matrix4 = ((SAPbouiCOM.Matrix)(this.GetItem("mochrg").Specific));
            this.Matrix4.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix4_KeyDownAfter);
            this.Matrix4.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix4_ClickAfter);
            this.Matrix4.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix4_ValidateAfter);
            this.Matrix4.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix4_LinkPressedBefore);
            this.Matrix5 = ((SAPbouiCOM.Matrix)(this.GetItem("mattach").Specific));
            this.Matrix5.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.Matrix5_DoubleClickAfter);
            this.Matrix5.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix5_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btnbrowse").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("btndisp").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("btndel").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tentry").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lbpcode").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("tmodcode").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("tmodname").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkmodcode").Specific));
            this.LinkedButton1.PressedBefore += new SAPbouiCOM._ILinkedButtonEvents_PressedBeforeEventHandler(this.LinkedButton1_PressedBefore);
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("tqteent").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("modelrow").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("lmodrow").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("lcompsum").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("tcompsum").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("lboutsum").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("tboutsum").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("ltprosum").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("ttprosum").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("lschrgsum").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("tschrgsum").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("lochrgsum").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("tochrgsum").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("ltotal").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("ttotal").Specific));
            this.Folder6 = ((SAPbouiCOM.Folder)(this.GetItem("fspare").Specific));
            this.Folder6.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder6_PressedAfter);
            this.Matrix6 = ((SAPbouiCOM.Matrix)(this.GetItem("mspare").Specific));
            this.Matrix6.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix6_KeyDownAfter);
            this.Matrix6.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix6_ClickAfter);
            this.Matrix6.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix6_ValidateAfter);
            this.Matrix6.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix6_LinkPressedBefore);
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("lsparsum").Specific));
            this.EditText17 = ((SAPbouiCOM.EditText)(this.GetItem("tsparsum").Specific));
            this.Folder7 = ((SAPbouiCOM.Folder)(this.GetItem("fccharge").Specific));
            this.Folder7.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder7_PressedAfter);
            this.Folder8 = ((SAPbouiCOM.Folder)(this.GetItem("fwcharge").Specific));
            this.Folder8.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder8_PressedAfter);
            this.Folder9 = ((SAPbouiCOM.Folder)(this.GetItem("fncharge").Specific));
            this.Folder9.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder9_PressedAfter);
            this.Matrix7 = ((SAPbouiCOM.Matrix)(this.GetItem("mccharge").Specific));
            this.Matrix7.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix7_KeyDownAfter);
            this.Matrix7.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix7_ClickAfter);
            this.Matrix7.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix7_LinkPressedBefore);
            this.Matrix7.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix7_ValidateAfter);
            this.Matrix8 = ((SAPbouiCOM.Matrix)(this.GetItem("mwcharge").Specific));
            this.Matrix8.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix8_KeyDownAfter);
            this.Matrix8.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix8_ClickAfter);
            this.Matrix8.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix8_LinkPressedBefore);
            this.Matrix8.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix8_ValidateAfter);
            this.Matrix9 = ((SAPbouiCOM.Matrix)(this.GetItem("mncharge").Specific));
            this.Matrix9.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix9_KeyDownAfter);
            this.Matrix9.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix9_ClickAfter);
            this.Matrix9.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix9_LinkPressedBefore);
            this.Matrix9.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix9_ValidateAfter);
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("lconsum").Specific));
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("tconsum").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("lwarr").Specific));
            this.EditText19 = ((SAPbouiCOM.EditText)(this.GetItem("twarr").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("lnego").Specific));
            this.EditText20 = ((SAPbouiCOM.EditText)(this.GetItem("tnego").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("lrtotal").Specific));
            this.EditText21 = ((SAPbouiCOM.EditText)(this.GetItem("trtotal").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chkround").Specific));
            this.CheckBox0.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.DataUpdateAfter += new SAPbouiCOM.Framework.FormBase.DataUpdateAfterHandler(this.Form_DataUpdateAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void OnCustomInitialize()
        {
            try
            {
                objform.Freeze(true);
                objform.Left = (clsModule.objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2;//form.Left + 50;
                objform.Top = (clsModule.objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4;// form.Top + 50;
                objform.ClientHeight = Button0.Item.Top + Button0.Item.Height + 10;
                //objform.ClientWidth = Folder5.Item.Left + Folder5.Item.Width + 10; 
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN");
                odbdsComponent = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN1"); //Components 
                odbdsBoughtOut = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN2"); //BoughtOut
                odbdsTestProp = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN3"); //Test Prop
                odbdsServiceChrge = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN4"); //Service Charges
                odbdsOverheadProp = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN5"); //Overhead Prop       
                odbdsAttachment = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN6"); //Attachments
                odbdsSpares = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN7"); //Spares
                odbdsContingency = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN8"); //Contingency              
                odbdsWarranty = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN9"); //Warranty             
                odbdsNegotiation = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN10"); //Negotiation   
                clsModule.objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_PCQTETRAN");
                ((SAPbouiCOM.EditText)objform.Items.Item("tdocdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "series", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocnum", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cmbstat", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocdate", true, true, false);         

                Folder5.Item.Left = Folder9.Item.Left + Folder9.Item.Width +1; //Attachments Tab Position                
                Folder6.Item.Left = Folder3.Item.Left + Folder3.Item.Width -1; //Spares Tab Position                
                Matrix5.Columns.Item("srcpath").Visible = false;
                Matrix5.Columns.Item("fileext").Visible = false;
                
                Folder0.Item.Click();
                objform.EnableMenu("1283", false); objform.EnableMenu("1284", false);
                objform.EnableMenu("1281", false); objform.EnableMenu("1282", false); //Find & Add
                objform.EnableMenu("1288", false); objform.EnableMenu("1289", false); objform.EnableMenu("1290", false); objform.EnableMenu("1291", false); //Next, Previous, First, Last Record
                Matrix0.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix1.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix2.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix3.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix4.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix6.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix7.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix8.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix9.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;               
                objform.EnableMenu("1304", true);//Refresh
                StaticText13.Item.TextStyle =Convert.ToInt16(SAPbouiCOM.BoTextStyle.ts_BOLD);
                StaticText18.Item.TextStyle =Convert.ToInt16(SAPbouiCOM.BoTextStyle.ts_BOLD);
                clsModule.objaddon.objGlobalVariables.bModal = true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
           
        }

        #region Fields        

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Folder Folder4;
        private SAPbouiCOM.Folder Folder5;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.Matrix Matrix4;
        private SAPbouiCOM.Matrix Matrix5;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.Folder Folder6;
        private SAPbouiCOM.Matrix Matrix6;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText17;
        private SAPbouiCOM.Folder Folder7;
        private SAPbouiCOM.Folder Folder8;
        private SAPbouiCOM.Folder Folder9;
        private SAPbouiCOM.Matrix Matrix7;
        private SAPbouiCOM.Matrix Matrix8;
        private SAPbouiCOM.Matrix Matrix9;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.EditText EditText19;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.EditText EditText20;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.EditText EditText21;
        private SAPbouiCOM.CheckBox CheckBox0;

        #endregion

        #region Components

        private void Folder0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mcomp";
                Matrix0.AutoResizeColumns();
                objform.Freeze(false);
                Matrix0.Columns.Item(1).Editable = false;
                Matrix0.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        } //Components

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "compcode":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR1\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_CompCode\"='" + odbdsComponent.GetValue("U_CompCode", pVal.Row - 1) + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmComponent component = new FrmComponent(((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        component.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "weight":
                    case "rate":
                        double total = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("weight").Cells.Item(pVal.Row).Specific).String) * Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("rate").Cells.Item(pVal.Row).Specific).String);
                        Matrix0.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(total));
                        Calculate_Total();
                        Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //try
            //{
            //    if (pVal.CharPressed == 9 || pVal.CharPressed == 37 || pVal.CharPressed == 38 || pVal.CharPressed == 39 || pVal.CharPressed == 40) return;
            //    if (pVal.Row >= 1 && (pVal.ColUID == "compcode" || pVal.ColUID == "compname"))
            //    {
            //        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR1\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_CompCode\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("compcode").Cells.Item(pVal.Row).Specific).String + "'");
            //        if (strSQL != "")
            //        {
            //            BubbleEvent = false;
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //}

        }

        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.ColUID == "total") BubbleEvent = false;
                int colID = Matrix0.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix0.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix0.VisualRowCount != pVal.Row) Matrix0.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix0.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix0.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix0.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix0.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix0_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //try
            //{
            //    strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR1\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_CompCode\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("compcode").Cells.Item(pVal.Row).Specific).String + "'"); //odbdsComponent.GetValue("U_CompCode", pVal.Row - 1)
            //    if (strSQL != "" && (pVal.ColUID == "compcode" || pVal.ColUID == "compname"))
            //    {
            //        objform.EnableMenu("771", false); objform.EnableMenu("773", false); objform.EnableMenu("774", false);
            //    }
            //    else objform.EnableMenu("773", true);
            //}
            //catch (Exception ex)
            //{
            //    //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //}
        }

        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "compcode" && pVal.ColUID != "compname") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR1\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_CompCode\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("compcode").Cells.Item(pVal.Row).Specific).String + "'"); //odbdsComponent.GetValue("U_CompCode", pVal.Row - 1)
                if (strSQL == "")
                {
                    Matrix0.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix0.Columns.Item(0).Editable = false; Matrix0.Columns.Item(Matrix0.Columns.Count-1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region BoughtOut

        private void Folder1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mbout";
                Matrix1.AutoResizeColumns();
                objform.Freeze(false);
                Matrix1.Columns.Item(1).Editable = false;
                Matrix1.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //BoughtOut

        private void Matrix1_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "boutc":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR2\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_BOutCode\"='" + ((SAPbouiCOM.EditText)Matrix1.Columns.Item("boutc").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmBoughtOut boughtOut = new FrmBoughtOut(((SAPbouiCOM.EditText)Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        boughtOut.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Matrix1_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "marper":
                    case "price":
                        decimal margintotal = Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix1.Columns.Item("price").Cells.Item(pVal.Row).Specific).String) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix1.Columns.Item("marper").Cells.Item(pVal.Row).Specific).String) / 100;
                        Matrix1.SetCellWithoutValidation(pVal.Row, "martot", Convert.ToString(margintotal));
                        margintotal = Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix1.Columns.Item("price").Cells.Item(pVal.Row).Specific).String) + Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix1.Columns.Item("martot").Cells.Item(pVal.Row).Specific).String);
                        Matrix1.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(margintotal));
                        Calculate_Total();
                        Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                   
                }
                objform.Freeze(true);
                Matrix1.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }       

        private void Matrix1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix1.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix1.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix1.VisualRowCount != pVal.Row) Matrix1.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix1.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix1.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix1.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix1.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "boutc" && pVal.ColUID != "boutn") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR2\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_BOutCode\"='" + ((SAPbouiCOM.EditText)Matrix1.Columns.Item("boutc").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix1.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix1.Columns.Item(0).Editable = false; Matrix1.Columns.Item(Matrix1.Columns.Count - 1).Editable = false; Matrix1.Columns.Item(Matrix1.Columns.Count - 2).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Testing Properties

        private void Folder2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mtprop";
                Matrix2.AutoResizeColumns();
                objform.Freeze(false);
                Matrix2.Columns.Item(1).Editable = false;
                Matrix2.Columns.Item(2).Editable = false;
                
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Test Properties

        private void Matrix2_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "tpropc":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR3\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_TPropCode\"='" + ((SAPbouiCOM.EditText)Matrix2.Columns.Item("tpropc").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmTestProperty testProperty = new FrmTestProperty(((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        testProperty.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix2_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "total":
                        Calculate_Total();
                        Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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

        private void Matrix2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "tpropc" && pVal.ColUID != "tpropn") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR3\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_TPropCode\"='" + ((SAPbouiCOM.EditText)Matrix2.Columns.Item("tpropc").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix2.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix2.Columns.Item(0).Editable = false; 
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }       

        private void Matrix2_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix2.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix2.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix2.VisualRowCount != pVal.Row) Matrix2.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix2.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix2.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix2.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix2.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #endregion

        #region Service Charges

        private void Folder3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mschrg";
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "servnam", "#");
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Service Charge

        private void Matrix3_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "servnam":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "servnam", "#");
                        break;
                    case "price":
                    case "qty":
                        decimal total = Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix3.Columns.Item("qty").Cells.Item(pVal.Row).Specific).String) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix3.Columns.Item("price").Cells.Item(pVal.Row).Specific).String);
                        Matrix3.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(total));
                        Calculate_Total();
                        Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                    //case "total":                       
                    //    break;
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

        private void Matrix3_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix3.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix3.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix3.VisualRowCount != pVal.Row) Matrix3.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix3.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix3.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix3.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix3.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Packing Charges

        private void Folder4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mochrg";
                Matrix4.AutoResizeColumns();                
                objform.Freeze(false);
                Matrix4.Columns.Item(1).Editable = false;
                Matrix4.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        } //Packing Charge   

        private void Matrix4_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix4.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "ochrgc":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR4\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_OPropCode\"='" + ((SAPbouiCOM.EditText)Matrix4.Columns.Item("ochrgc").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmPackingCharges overheadProperty = new FrmPackingCharges(((SAPbouiCOM.EditText)Matrix4.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        overheadProperty.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix4_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "marper":
                    case "total":
                        if (((SAPbouiCOM.EditText)Matrix4.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                        decimal margintotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0))) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix4.Columns.Item("marper").Cells.Item(pVal.Row).Specific).String) / 100;
                        Matrix4.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(margintotal));
                        Calculate_Total();
                        Matrix4.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                }
                objform.Freeze(true);
                Matrix4.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }
        }        

        private void Matrix4_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix4.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix4.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix4.VisualRowCount != pVal.Row) Matrix4.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix4.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix4.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix4.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix4.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "ochrgc" && pVal.ColUID != "ochrgn") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR4\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_OPropCode\"='" + ((SAPbouiCOM.EditText)Matrix4.Columns.Item("ochrgc").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix4.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix4.Columns.Item(0).Editable = false; Matrix4.Columns.Item(Matrix4.Columns.Count - 1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Spares

        private void Folder6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mspare";
                Matrix6.AutoResizeColumns();
                objform.Freeze(false);
                Matrix6.Columns.Item(1).Editable = false;
                Matrix6.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }        

        private void Matrix6_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix6.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix6.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix6.VisualRowCount != pVal.Row) Matrix6.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix6.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix6.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix6.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix6.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }       

        private void Matrix6_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "sparc":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR6\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_SparCode\"='" + ((SAPbouiCOM.EditText)Matrix6.Columns.Item("sparc").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmSpareMaster spareMaster = new FrmSpareMaster(((SAPbouiCOM.EditText)Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        spareMaster.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        private void Matrix6_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "total":
                        if (pVal.ItemChanged == false) return;
                        Calculate_Total();
                        Matrix6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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

        private void Matrix6_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "sparc" && pVal.ColUID != "sparc") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR6\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_SparCode\"='" + ((SAPbouiCOM.EditText)Matrix6.Columns.Item("sparc").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix6.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix6.Columns.Item(0).Editable = false; Matrix6.Columns.Item(Matrix6.Columns.Count - 1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Contingency

        private void Folder7_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mccharge";
                Matrix7.AutoResizeColumns();
                objform.Freeze(false);
                Matrix7.Columns.Item(1).Editable = false;
                Matrix7.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix7_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "marper":
                    case "total":
                        if (pVal.ItemChanged == false) return;
                        decimal margintotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0))) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix7.Columns.Item("marper").Cells.Item(pVal.Row).Specific).String) / 100;
                        Matrix7.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(margintotal));
                        Calculate_Total();
                        Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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

        private void Matrix7_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix7.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix7.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix7.VisualRowCount != pVal.Row) Matrix7.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix7.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix7.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix7.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix7.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix7_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "contc":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR7\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_ContCode\"='" + ((SAPbouiCOM.EditText)Matrix7.Columns.Item("contc").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmContingencyCharges contingencyCharges = new FrmContingencyCharges(((SAPbouiCOM.EditText)Matrix7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        contingencyCharges.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Matrix7_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "contc" && pVal.ColUID != "contn") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR7\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_ContCode\"='" + ((SAPbouiCOM.EditText)Matrix7.Columns.Item("contc").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix7.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix7.Columns.Item(0).Editable = false; Matrix7.Columns.Item(Matrix7.Columns.Count - 1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Warranty

        private void Folder8_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mwcharge";
                Matrix8.AutoResizeColumns();
                objform.Freeze(false);
                Matrix8.Columns.Item(1).Editable = false;
                Matrix8.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix8_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "marper":
                    case "total":
                        if (pVal.ItemChanged == false) return;
                        decimal margintotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_ContSum", 0))) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix8.Columns.Item("marper").Cells.Item(pVal.Row).Specific).String) / 100;
                        Matrix8.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(margintotal));
                        Calculate_Total();
                        Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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

        private void Matrix8_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix8.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix8.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix8.VisualRowCount != pVal.Row) Matrix8.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix8.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix8.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix8.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix8.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix8_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "wcode":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR8\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_WarrCode\"='" + ((SAPbouiCOM.EditText)Matrix8.Columns.Item("wcode").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmWarrantyCharges warrantyCharges = new FrmWarrantyCharges(((SAPbouiCOM.EditText)Matrix8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        warrantyCharges.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Matrix8_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "wcode" && pVal.ColUID != "wname") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR8\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_WarrCode\"='" + ((SAPbouiCOM.EditText)Matrix8.Columns.Item("wcode").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix8.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix8.Columns.Item(0).Editable = false; Matrix8.Columns.Item(Matrix8.Columns.Count - 1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Negotiation

        private void Folder9_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mncharge";
                Matrix9.AutoResizeColumns();
                objform.Freeze(false);
                Matrix9.Columns.Item(1).Editable = false;
                Matrix9.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }

        private void Matrix9_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (((SAPbouiCOM.EditText)Matrix9.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "" || pVal.ItemChanged == false) return;
                switch (pVal.ColUID)
                {
                    case "marper":
                    case "total":
                        if (pVal.ItemChanged == false) return;
                        decimal margintotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_ContSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_WarrSum", 0))) * Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix9.Columns.Item("marper").Cells.Item(pVal.Row).Specific).String) / 100;
                        Matrix9.SetCellWithoutValidation(pVal.Row, "total", Convert.ToString(margintotal));
                        Calculate_Total();
                        Matrix9.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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

        private void Matrix9_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int colID = Matrix9.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    if (pVal.Row > 1) Matrix9.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    if (Matrix9.VisualRowCount != pVal.Row) Matrix9.SetCellFocus(pVal.Row + 1, colID);
                }
                //else if (pVal.CharPressed == 37 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Left
                //{
                //    if (Matrix9.Columns.Item(colID - 1).Editable == false) return;
                //    Matrix9.SetCellFocus(pVal.Row, colID - 1);
                //}
                //else if (pVal.CharPressed == 39 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Right
                //{
                //    if (Matrix9.Columns.Item(colID + 1).Editable == false) return;
                //    Matrix9.SetCellFocus(pVal.Row, colID + 1);
                //}
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Matrix9_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (((SAPbouiCOM.EditText)Matrix9.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String == "") return;
                switch (pVal.ColUID)
                {
                    case "ngcode":
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR9\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_NegoCode\"='" + ((SAPbouiCOM.EditText)Matrix9.Columns.Item("ngcode").Cells.Item(pVal.Row).Specific).String + "'");
                        if (strSQL == "") { clsModule.objaddon.objapplication.StatusBar.SetText("Entry is not found in the master... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                        FrmNegotiationCharges negotiationCharges = new FrmNegotiationCharges(((SAPbouiCOM.EditText)Matrix9.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String);
                        negotiationCharges.Show();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Matrix9_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID != "ngcode" && pVal.ColUID != "ngname") return;
                strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select distinct 1 \"Status\" from \"@AT_PCQTMSTR9\" T0 Where T0.\"Code\"='" + odbdsHeader.GetValue("U_ModelCode", 0) + "' and T0.\"U_NegoCode\"='" + ((SAPbouiCOM.EditText)Matrix9.Columns.Item("ngcode").Cells.Item(pVal.Row).Specific).String + "'");
                if (strSQL == "")
                {
                    Matrix9.CommonSetting.SetRowEditable(pVal.Row, true);
                    Matrix9.Columns.Item(0).Editable = false; Matrix9.Columns.Item(Matrix9.Columns.Count - 1).Editable = false;
                }
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Header Fields & Events

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                if (EditText5.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("PreCost Quote Selection is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText1.Item.Click();
                    return;
                }
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix0, "compcode");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix1, "boutc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix2, "tpropc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix3, "servnam");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix4, "ochrgc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix6, "sparc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix7, "contc");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix8, "wcode");
                clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix9, "ngcode");
                //if (clsModule.objaddon.objapplication.MessageBox("You cannot change this document after you have added it. Continue?", 2, "Yes", "No") != 1) { BubbleEvent = false; return; }

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == true && objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    objform.Close();
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("PCQTETRAN", pVal.FormTypeCount);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        
        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                if (pVal.ActionSuccess == true)
                {
                    strSQL = odbdsHeader.GetValue("DocEntry", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("quotetran").Cells.Item(selectionQteLine).Specific).String = strSQL;
                    strSQL = odbdsHeader.GetValue("U_RTotal", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("total").Cells.Item(selectionQteLine).Specific).String = strSQL;
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("compsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_CompSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("boutsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_BoutSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("tpropsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_TPropSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("schrgsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_SChrgSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("sparsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_SparSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("packsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_OPropSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("contsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_ContSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("warsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_WarrSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("negosum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_NegoSum", 0);
                    if (selectionForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) selectionForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    selectionForm.Items.Item("1").Click();
                    selectionForm = null;
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
               //if(objform.Visible) objform.ClientHeight = Button0.Item.Top + Button0.Item.Height + 10;
               //if(objform.Visible) objform.ClientWidth = objform.MaxWidth;
                if (Folder0.Selected == true) Matrix0.AutoResizeColumns();
                else if (Folder1.Selected == true) Matrix1.AutoResizeColumns();
                else if (Folder2.Selected == true) Matrix2.AutoResizeColumns();
                else if (Folder3.Selected == true) Matrix3.AutoResizeColumns();
                else if (Folder4.Selected == true) Matrix4.AutoResizeColumns();
                else if (Folder5.Selected == true) Matrix5.AutoResizeColumns();
                else if (Folder6.Selected == true) Matrix6.AutoResizeColumns();
                else if (Folder7.Selected == true) Matrix7.AutoResizeColumns();
                else if (Folder8.Selected == true) Matrix8.AutoResizeColumns();
                else if (Folder9.Selected == true) Matrix9.AutoResizeColumns();
                
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }               

        private void Form_DataUpdateAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                if (pVal.ActionSuccess == true)
                {
                    strSQL = odbdsHeader.GetValue("U_RTotal", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("total").Cells.Item(selectionQteLine).Specific).String = strSQL;
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("compsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_CompSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("boutsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_BoutSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("tpropsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_TPropSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("schrgsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_SChrgSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("sparsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_SparSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("packsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_OPropSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("contsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_ContSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("warsum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_WarrSum", 0);
                    ((SAPbouiCOM.EditText)((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("negosum").Cells.Item(selectionQteLine).Specific).String = odbdsHeader.GetValue("U_NegoSum", 0);
                    if (selectionForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) selectionForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    selectionForm.Items.Item("1").Click();
                    ((SAPbouiCOM.Matrix)selectionForm.Items.Item("mprodmod").Specific).Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                }
            }
            catch (Exception ex)
            {
            }
        }               

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Refreshing details. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                LoadDataByQuery(odbdsHeader.GetValue("U_ModelCode", 0), odbdsHeader.GetValue("DocEntry", 0));
                clsModule.objaddon.objapplication.StatusBar.SetText("Refreshed successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
            }
        }

        private void LinkedButton1_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                FrmPreCostQuoteMaster preCostQuoteMaster = new FrmPreCostQuoteMaster(EditText7.Value);
                preCostQuoteMaster.Show();
            }
            catch (Exception ex)
            {

            }

        }

        private void CheckBox0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                if (CheckBox0.Checked == true)
                {
                    strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select case when Right(Cast(" + odbdsHeader.GetValue("U_Total", 0) + " as Integer),2)>50 then ROUND(" + odbdsHeader.GetValue("U_Total", 0) + ",-2) Else ROUND(" + odbdsHeader.GetValue("U_Total", 0) + ",-2) + 100 End from dummy");
                    odbdsHeader.SetValue("U_RTotal", 0, strSQL);
                }
                else
                {
                    odbdsHeader.SetValue("U_RTotal", 0, Convert.ToString(odbdsHeader.GetValue("U_Total", 0)));
                }
            }
            catch (Exception ex)
            {

            }
        }

        #endregion

        #region Attachments      

        private void Folder5_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mattach";
                Matrix5.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Attachments 

        private void Matrix5_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Matrix5.SelectRow(pVal.Row, true, false);
                if (Matrix5.IsRowSelected(pVal.Row) == true)
                {
                    objform.Items.Item("btndisp").Enabled = true;
                    objform.Items.Item("btndel").Enabled = true;
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void Matrix5_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix5, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {

            }

        }       

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.SetAttachMentFile(objform, odbdsHeader, Matrix5, odbdsAttachment);
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

        }

        private void Button3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix5, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.DeleteRowAttachment(objform, Matrix5, odbdsAttachment, Matrix5.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }  
        
        #endregion

        #region Functions

        public void LoadDataByQuery(string ModelCode,string DocEntry)
        {
            try
            {
                bool flag = false;
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                
                strSQL = "Select T0.\"U_CompCode\",T0.\"U_CompName\",T0.\"U_Weight\" from \"@AT_PCQTMSTR1\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_CompCode\"<>''";
                if(DocEntry!="") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN1\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    //Matrix0.Clear();
                    odbdsComponent.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix0.AddRow();
                        Matrix0.GetLineData(Matrix0.VisualRowCount);
                        odbdsComponent.SetValue("LineId", 0, Convert.ToString(Matrix0.VisualRowCount));
                        odbdsComponent.SetValue("U_CompCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsComponent.SetValue("U_CompName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        odbdsComponent.SetValue("U_Weight", 0, Convert.ToString(objrs.Fields.Item(2).Value));
                        Matrix0.SetLineData(Matrix0.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_BOutCode\",T0.\"U_BOutName\" from \"@AT_PCQTMSTR2\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_BOutCode\"<>''";
                if (DocEntry != "")  strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN2\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    //Matrix1.Clear();
                    odbdsBoughtOut.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix1.AddRow();
                        Matrix1.GetLineData(Matrix1.VisualRowCount);
                        odbdsBoughtOut.SetValue("LineId", 0, Convert.ToString(Matrix1.VisualRowCount));
                        odbdsBoughtOut.SetValue("U_BOutCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsBoughtOut.SetValue("U_BOutName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix1.SetLineData(Matrix1.VisualRowCount);
                        objrs.MoveNext();
                    }
                }                    
        
                strSQL = "Select T0.\"U_TPropCode\",T0.\"U_TPropName\" from \"@AT_PCQTMSTR3\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_TPropCode\"<>''";
                if (DocEntry != "")  strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN3\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    //Matrix2.Clear();
                    odbdsTestProp.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix2.AddRow();
                        Matrix2.GetLineData(Matrix2.VisualRowCount);
                        odbdsTestProp.SetValue("LineId", 0, Convert.ToString(Matrix2.VisualRowCount));
                        odbdsTestProp.SetValue("U_TPropCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsTestProp.SetValue("U_TPropName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix2.SetLineData(Matrix2.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_OPropCode\",T0.\"U_OPropName\" from \"@AT_PCQTMSTR4\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_OPropCode\"<>''";
                if (DocEntry != "") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN5\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    //Matrix4.Clear();
                    odbdsOverheadProp.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix4.AddRow();
                        Matrix4.GetLineData(Matrix4.VisualRowCount);
                        odbdsOverheadProp.SetValue("LineId", 0, Convert.ToString(Matrix4.VisualRowCount));
                        odbdsOverheadProp.SetValue("U_OPropCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsOverheadProp.SetValue("U_OPropName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix4.SetLineData(Matrix4.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_SparCode\",T0.\"U_SparName\" from \"@AT_PCQTMSTR6\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_SparCode\"<>''";
                if (DocEntry != "") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN7\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    odbdsSpares.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix6.AddRow();
                        Matrix6.GetLineData(Matrix6.VisualRowCount);
                        odbdsSpares.SetValue("LineId", 0, Convert.ToString(Matrix6.VisualRowCount));
                        odbdsSpares.SetValue("U_SparCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsSpares.SetValue("U_SparName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix6.SetLineData(Matrix6.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_ContCode\",T0.\"U_ContName\" from \"@AT_PCQTMSTR7\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_ContCode\"<>''";
                if (DocEntry != "") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN8\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    odbdsContingency.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix7.AddRow();
                        Matrix7.GetLineData(Matrix7.VisualRowCount);
                        odbdsContingency.SetValue("LineId", 0, Convert.ToString(Matrix7.VisualRowCount));
                        odbdsContingency.SetValue("U_ContCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsContingency.SetValue("U_ContName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix7.SetLineData(Matrix7.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_WarrCode\",T0.\"U_WarrName\" from \"@AT_PCQTMSTR8\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_WarrCode\"<>''";
                if (DocEntry != "") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN9\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    odbdsWarranty.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix8.AddRow();
                        Matrix8.GetLineData(Matrix8.VisualRowCount);
                        odbdsWarranty.SetValue("LineId", 0, Convert.ToString(Matrix8.VisualRowCount));
                        odbdsWarranty.SetValue("U_WarrCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsWarranty.SetValue("U_WarrName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix8.SetLineData(Matrix8.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                strSQL = "Select T0.\"U_NegoCode\",T0.\"U_NegoName\" from \"@AT_PCQTMSTR9\" T0 Where T0.\"Code\"='" + ModelCode + "' and T0.\"U_NegoCode\"<>''";
                if (DocEntry != "") strSQL += "\n and T0.\"LineId\" > ifnull((Select max(\"LineId\") from \"@AT_PCQTETRAN10\" Where \"DocEntry\"='" + DocEntry + "'),0)";
                objrs.DoQuery(strSQL);
                if (objrs.RecordCount > 0)
                {
                    odbdsNegotiation.Clear();
                    while (!objrs.EoF)
                    {
                        flag = true;
                        Matrix9.AddRow();
                        Matrix9.GetLineData(Matrix9.VisualRowCount);
                        odbdsNegotiation.SetValue("LineId", 0, Convert.ToString(Matrix9.VisualRowCount));
                        odbdsNegotiation.SetValue("U_NegoCode", 0, Convert.ToString(objrs.Fields.Item(0).Value));
                        odbdsNegotiation.SetValue("U_NegoName", 0, Convert.ToString(objrs.Fields.Item(1).Value));
                        Matrix9.SetLineData(Matrix9.VisualRowCount);
                        objrs.MoveNext();
                    }
                }

                if (flag)
                {
                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;                    
                }
                Field_Editable_SetUp(Matrix0);
                Field_Editable_SetUp(Matrix1);
                Field_Editable_SetUp(Matrix2);
                Field_Editable_SetUp(Matrix4);
                Field_Editable_SetUp(Matrix6);
                Field_Editable_SetUp(Matrix7);
                Field_Editable_SetUp(Matrix8);
                Field_Editable_SetUp(Matrix9);
            }
            catch (Exception ex)
            {
            }
        }

        private void Calculate_Total()
        {
            try
            {
                decimal total=0, overAllTotal=0;
                objform.Freeze(true);
                if (Folder0.Selected==true)
                {
                    Matrix0.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsComponent.Size - 1; Row++)
                    {
                        if (odbdsComponent.GetValue("U_CompCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsComponent.GetValue("U_Total", Row));
                        }
                    }
                    Matrix0.LoadFromDataSource();
                    odbdsHeader.SetValue("U_CompSum", 0, Convert.ToString(total));                    
                }
                else if (Folder1.Selected == true)
                {
                    Matrix1.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsBoughtOut.Size - 1; Row++)
                    {
                        if (odbdsBoughtOut.GetValue("U_BOutCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsBoughtOut.GetValue("U_Total", Row));
                        }
                    }
                    Matrix1.LoadFromDataSource();
                    odbdsHeader.SetValue("U_BoutSum", 0, Convert.ToString(total));
                }
                else if (Folder2.Selected == true)
                {
                    Matrix2.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsTestProp.Size - 1; Row++)
                    {
                        if (odbdsTestProp.GetValue("U_TPropCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsTestProp.GetValue("U_Total", Row));
                        }
                    }
                    Matrix2.LoadFromDataSource();
                    odbdsHeader.SetValue("U_TPropSum", 0, Convert.ToString(total));
                }
                else if (Folder3.Selected == true)
                {
                    Matrix3.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsServiceChrge.Size - 1; Row++)
                    {
                        if (odbdsServiceChrge.GetValue("U_ServName", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsServiceChrge.GetValue("U_Total", Row));
                        }
                    }
                    Matrix3.LoadFromDataSource();
                    odbdsHeader.SetValue("U_SChrgSum", 0, Convert.ToString(total));
                }
                else if (Folder4.Selected == true)
                {
                    Matrix4.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsOverheadProp.Size - 1; Row++)
                    {
                        if (odbdsOverheadProp.GetValue("U_OPropCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsOverheadProp.GetValue("U_Total", Row));
                        }                        
                    }
                    Matrix4.LoadFromDataSource();
                    odbdsHeader.SetValue("U_OPropSum", 0, Convert.ToString(total));                   

                }
                else if (Folder6.Selected == true)
                {
                    Matrix6.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsSpares.Size - 1; Row++)
                    {
                        if (odbdsSpares.GetValue("U_SparCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsSpares.GetValue("U_Total", Row));
                        }
                    }
                    Matrix6.LoadFromDataSource();
                    odbdsHeader.SetValue("U_SparSum", 0, Convert.ToString(total));
                }
                else if (Folder7.Selected == true)
                {
                    Matrix7.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsContingency.Size - 1; Row++)
                    {
                        if (odbdsContingency.GetValue("U_ContCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsContingency.GetValue("U_Total", Row));
                        }                        
                    }
                    Matrix7.LoadFromDataSource();
                    odbdsHeader.SetValue("U_ContSum", 0, Convert.ToString(total));
                }
                else if (Folder8.Selected == true)
                {
                    Matrix8.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsWarranty.Size - 1; Row++)
                    {
                        if (odbdsWarranty.GetValue("U_WarrCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsWarranty.GetValue("U_Total", Row));
                        }
                    }
                    Matrix8.LoadFromDataSource();
                    odbdsHeader.SetValue("U_WarrSum", 0, Convert.ToString(total));
                }
                else if (Folder9.Selected == true)
                {
                    Matrix9.FlushToDataSource();
                    for (int Row = 0; Row <= odbdsNegotiation.Size - 1; Row++)
                    {
                        if (odbdsNegotiation.GetValue("U_NegoCode", Row) != "")
                        {
                            total += Convert.ToDecimal(odbdsNegotiation.GetValue("U_Total", Row));
                        }
                    }
                    Matrix9.LoadFromDataSource();
                    odbdsHeader.SetValue("U_NegoSum", 0, Convert.ToString(total));
                }
                Calculate_MarginTotal();
                overAllTotal += Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0))+ Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0))+ Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0))+ Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_ContSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_WarrSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_NegoSum", 0));
                odbdsHeader.SetValue("U_Total", 0, Convert.ToString(overAllTotal));

                if (odbdsHeader.GetValue("U_Rounding", 0) == "Y")
                {
                    strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select case when Right(Cast("+ overAllTotal + " as Integer),2)>50 then ROUND(" + overAllTotal + ",-2) Else ROUND(" + overAllTotal + ",-2) + 100 End from dummy");
                    odbdsHeader.SetValue("U_RTotal", 0, Convert.ToString(strSQL));
                }
                else
                {
                    odbdsHeader.SetValue("U_RTotal", 0, Convert.ToString(overAllTotal));
                }

            }
            catch (Exception ex)
            {
            }
            finally
            {
                objform.Freeze(false);
            }
        }

        private void Field_Editable_SetUp(SAPbouiCOM. Matrix matrix)
        {
            try
            {
                if (matrix.Columns.Item(1).Editable == false && matrix.Columns.Item(2).Editable == false) return;
                matrix.Columns.Item(1).Editable = false;
                matrix.Columns.Item(2).Editable = false;
            }
            catch (Exception ex)
            {

            }
        }

        private void Calculate_MarginTotal()
        {
            try
            {
                decimal marginTotal = 0, total=0;
                Matrix4.FlushToDataSource();
                for (int Row = 0; Row <= odbdsOverheadProp.Size - 1; Row++)
                {
                    if (odbdsOverheadProp.GetValue("U_OPropCode", Row) != "")
                    {                        
                        marginTotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0))) * Convert.ToDecimal(odbdsOverheadProp.GetValue("U_MarginPer", Row)) / 100;
                        odbdsOverheadProp.SetValue("U_Total", Row, Convert.ToString(marginTotal));
                        total += Convert.ToDecimal(odbdsOverheadProp.GetValue("U_Total", Row));
                    }
                }
                Matrix4.LoadFromDataSource();
                odbdsHeader.SetValue("U_OPropSum", 0, Convert.ToString(total));
                total = 0;
                Matrix7.FlushToDataSource();
                for (int Row = 0; Row <= odbdsContingency.Size - 1; Row++)
                {
                    if (odbdsContingency.GetValue("U_ContCode", Row) != "")
                    {                        
                        marginTotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0))) * Convert.ToDecimal(odbdsContingency.GetValue("U_MarginPer", Row)) / 100;
                        odbdsContingency.SetValue("U_Total", Row, Convert.ToString(marginTotal));
                        total += Convert.ToDecimal(odbdsContingency.GetValue("U_Total", Row));
                    }
                }
                Matrix7.LoadFromDataSource();
                odbdsHeader.SetValue("U_ContSum", 0, Convert.ToString(total));
                total = 0;
                Matrix8.FlushToDataSource();
                for (int Row = 0; Row <= odbdsWarranty.Size - 1; Row++)
                {
                    if (odbdsWarranty.GetValue("U_WarrCode", Row) != "")
                    {                        
                        marginTotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_ContSum", 0))) * Convert.ToDecimal(odbdsWarranty.GetValue("U_MarginPer", Row)) / 100;
                        odbdsWarranty.SetValue("U_Total", Row, Convert.ToString(marginTotal));
                        total += Convert.ToDecimal(odbdsWarranty.GetValue("U_Total", Row));
                    }
                }
                Matrix8.LoadFromDataSource();
                odbdsHeader.SetValue("U_WarrSum", 0, Convert.ToString(total));
                total = 0;
                Matrix9.FlushToDataSource();
                for (int Row = 0; Row <= odbdsNegotiation.Size - 1; Row++)
                {
                    if (odbdsNegotiation.GetValue("U_NegoCode", Row) != "")
                    {                        
                        marginTotal = (Convert.ToDecimal(odbdsHeader.GetValue("U_CompSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_BoutSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_TPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SChrgSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_SparSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_OPropSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_ContSum", 0)) + Convert.ToDecimal(odbdsHeader.GetValue("U_WarrSum", 0))) * Convert.ToDecimal(odbdsNegotiation.GetValue("U_MarginPer", Row)) / 100;
                        odbdsNegotiation.SetValue("U_Total", Row, Convert.ToString(marginTotal));
                        total += Convert.ToDecimal(odbdsNegotiation.GetValue("U_Total", Row));
                    }
                }
                Matrix9.LoadFromDataSource();
                odbdsHeader.SetValue("U_NegoSum", 0, Convert.ToString(total));

            }
            catch (Exception ex)
            {

            }
        }

        #endregion

        
    }
}

