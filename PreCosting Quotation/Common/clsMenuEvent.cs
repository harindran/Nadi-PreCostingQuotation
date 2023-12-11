using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PreCosting_Quotation.Common
{
    class clsMenuEvent
    {     

        SAPbouiCOM.Form objform;
        string strsql;
        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule. objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "-392":
                    case "-393":
                    case "392":
                    case "393":
                        {
                            // Default_Sample_MenuEvent(pVal, BubbleEvent)
                            if (pVal.BeforeAction == true)
                                return;
                            objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            Default_Sample_MenuEvent(pVal, BubbleEvent);

                            break;
                        }
                    case "PCQUOTE":
                        PreCostQuote_Selection_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                    case "PCQTETRAN":
                        PreCostQuote_Transaction_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                    case "PCQTMSTR":
                        PreCost_QuoteMaster_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                }
            } 
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }

                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find
                            {
                                //oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                //oUDFForm.Items.Item("U_RevRecDE").Enabled = true;
                                break;
                            }
                        case "1287":
                            {
                                //if (oUDFForm.Items.Item("U_RevRecDN").Enabled == false|| oUDFForm.Items.Item("U_RevRecDE").Enabled == false)
                                //{
                                //    oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                //    oUDFForm.Items.Item("U_RevRecDE").Enabled = true;
                                //}
                                //((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_RevRecDN").Specific).String = "";
                                //((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_RevRecDE").Specific).String = "";
                                break;
                            }
                        default:
                            {
                                //oUDFForm.Items.Item("U_RevRecDN").Enabled = false;
                                //oUDFForm.Items.Item("U_RevRecDE").Enabled = false;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }

        private void PreCostQuote_Selection_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                SAPbobsCOM.Recordset objRs;
                SAPbouiCOM.DBDataSource DBSource;
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                DBSource = objform.DataSources.DBDataSources.Item("@AT_PCQTESEL");
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mprodmod").Specific;
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1284": //Cancel
                            if (clsModule.objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }                           
                            
                            break;
                        case "1286":
                            {
                                //clsModule.objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                //BubbleEvent = false;
                                //return;
                                break;
                            }
                        case "1293":
                            if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                            break;
                    }
                }
                else
                {                    
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                           
                            
                            objform.Items.Item("mprodmod").Enabled = false;
                            //objform.EnableMenu("1282", true);
                            objform.ActiveItem = "tdocnum";
                            break;
                            
                        case "1282"://Add Mode          
                            ((SAPbouiCOM.EditText)objform.Items.Item("tdocdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                            clsModule.objaddon.objglobalmethods.LoadSeries(objform, DBSource, "AT_PCQTESEL");                                
                                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                            clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "modcode", "#");
                            objform.EnableMenu("1282", false);
                            break;
                        case "1292"://Add Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fprodmod").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mprodmod").Specific, "modcode", "#");
                            break;
                                              
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void PreCostQuote_Transaction_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                //SAPbobsCOM.Recordset objRs;
                SAPbouiCOM.DBDataSource DBSource;
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                DBSource = objform.DataSources.DBDataSources.Item("@AT_PCQTETRAN");
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific;
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1283": //Remove
                        case "1284": //Cancel
                            if (clsModule.objaddon.objapplication.MessageBox("Removing or Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            break;
                        case "1292": //Add Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fcomp").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific, "compcode", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fbout").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mbout").Specific, "boutc", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mbout").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("ftprop").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mtprop").Specific, "tpropc", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mtprop").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fschrg").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mschrg").Specific, "servnam", "#"); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fochrg").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mochrg").Specific, "ochrgc", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mochrg").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fspare").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mspare").Specific, "sparc", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mspare").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fccharge").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mccharge").Specific, "contc", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mccharge").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fwcharge").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mwcharge").Specific, "wcode", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mwcharge").Specific); }
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fncharge").Specific).Selected == true) { clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mncharge").Specific, "ngcode", "#"); clsModule.objaddon.objglobalmethods.SetCellEditable((SAPbouiCOM.Matrix)objform.Items.Item("mncharge").Specific); }

                            break;
                        case "1293":
                            if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                            break;
                    }
                }
                else
                {
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                           
                            break;
                        case "1282"://Add Mode                              
                            break;
                        case "1293"://Delete Row

                            break;

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void PreCost_QuoteMaster_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                //SAPbouiCOM.DBDataSource  odbdsComponent, odbdsBoughtOut, odbdsTestProp, odbdsOverheadProp;
                //SAPbouiCOM.Matrix matComponent, matBoqItem,matBoqLabour;                
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                //Content Matrix
                //odbdsComponent = objform.DataSources.DBDataSources.Item("@AT_PCQTMSTR1");
                //matComponent = (SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific;                              

                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1283":
                            if (clsModule.objaddon.objapplication.MessageBox("Removing of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            break;
                        case "1292": //Add Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fcomp").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific, "compcode", "#"); 
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fbout").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mbout").Specific, "bocode", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("ftprop").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mtprop").Specific, "testproc", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("foprop").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("moprop").Specific, "opropc", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fspare").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mspare").Specific, "sparc", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fccharge").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mccharge").Specific, "contc", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fwcharge").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mwcharge").Specific, "wcode", "#");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fncharge").Specific).Selected == true) clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mncharge").Specific, "ngcode", "#");

                            break;
                        case "1293": //Delete Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fcomp").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fbout").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mbout").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("ftprop").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mtprop").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("foprop").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("moprop").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fspare").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mspare").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fccharge").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mccharge").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fwcharge").Specific).Selected == true)
                                if (((SAPbouiCOM.Matrix)objform.Items.Item("mwcharge").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fncharge").Specific).Selected == true)
                                 if (((SAPbouiCOM.Matrix)objform.Items.Item("mncharge").Specific).VisualRowCount == 1) { BubbleEvent = false; return; }

                            break;
                    }
                }
                else
                {
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                            
                            objform.Items.Item("mcomp").Enabled = false;
                            objform.Items.Item("mbout").Enabled = false;
                            objform.Items.Item("mtprop").Enabled = false;
                            objform.Items.Item("moprop").Enabled = false;
                            objform.Items.Item("mattach").Enabled = false;
                            objform.Items.Item("mspare").Enabled = false;
                            objform.Items.Item("mccharge").Enabled = false;
                            objform.Items.Item("mwcharge").Enabled = false;
                            objform.Items.Item("mncharge").Enabled = false;
                            break;
                        case "1293"://Delete Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fcomp").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific, "@AT_PCQTMSTR1");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fbout").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mbout").Specific, "@AT_PCQTMSTR2");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("ftprop").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mtprop").Specific, "@AT_PCQTMSTR3");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("foprop").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("moprop").Specific, "@AT_PCQTMSTR4");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fspare").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mspare").Specific, "@AT_PCQTMSTR6");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fccharge").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mccharge").Specific, "@AT_PCQTMSTR7");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fwcharge").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mwcharge").Specific, "@AT_PCQTMSTR8");
                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fncharge").Specific).Selected == true) DeleteRow((SAPbouiCOM.Matrix)objform.Items.Item("mncharge").Specific, "@AT_PCQTMSTR9");

                            break;
                        case "1282"://Add Mode                            
                            ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                            objform.Items.Item("fcomp").Click();
                            //clsModule.objaddon.objglobalmethods.Matrix_Addrow((SAPbouiCOM.Matrix)objform.Items.Item("mcomp").Specific, "compcode", "#");
                            objform.ActiveItem = "tfanmod";
                            break;                       

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void DeleteRow(SAPbouiCOM.Matrix objMatrix, string TableName)
        {
            try
            {
                SAPbouiCOM.DBDataSource DBSource;
                // objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource();
                DBSource = objform.DataSources.DBDataSources.Item(TableName); 
                for (int i = 1, loopTo = objMatrix.VisualRowCount; i <= loopTo; i++)
                {
                    objMatrix.GetLineData(i);
                    DBSource.Offset = i - 1;
                    DBSource.SetValue("LineId", DBSource.Offset, Convert.ToString(i));
                    objMatrix.SetLineData(i);
                    objMatrix.FlushToDataSource();
                }
                DBSource.RemoveRecord(DBSource.Size - 1);
                objMatrix.LoadFromDataSource();
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            finally
            {
            }
        }
       

    }
}
