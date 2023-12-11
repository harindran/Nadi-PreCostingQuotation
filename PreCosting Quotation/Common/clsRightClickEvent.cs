using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PreCosting_Quotation.Common
{
    class clsRightClickEvent
    {
        SAPbouiCOM.Form objform;
        clsGlobalMethods objglobalMethods= new clsGlobalMethods();
        SAPbouiCOM.ComboBox ocombo;
        SAPbouiCOM.Matrix Matrix0;
        string strsql;
        SAPbobsCOM.Recordset objrs;

        public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "PCMODEL": //Fan Model
                    case "PCCOMP": //Component
                    case "PCBOUT": //Boughtout
                    case "PCTPROP": //Test Property
                    case "PCOPROP": //Packing Charges
                    case "PCSPAR": //Spare master
                    case "PCCCGE": //Contingency Charges 
                    case "PCNCGE": //Negotiation Charges 
                    case "PCWCGE": //Warranty Charges 
                        if (eventInfo.BeforeAction == true)
                        {
                            objform.EnableMenu("1283", false);
                            if (eventInfo.ItemUID == "") { BubbleEvent = false; return; }
                            if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT || objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                            {
                                if (((SAPbouiCOM.EditText)objform.Items.Item(eventInfo.ItemUID).Specific).String != "") objform.EnableMenu("772", true);  // Copy
                                else objform.EnableMenu("773", true);  // Paste
                            }
                            else if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            {
                                if (((SAPbouiCOM.ComboBox)objform.Items.Item(eventInfo.ItemUID).Specific).Selected.Value != "") objform.EnableMenu("772", true);  // Copy
                                else objform.EnableMenu("773", true);  // Paste
                            }                            
                            else
                            {   
                                if (eventInfo.ItemUID != "") objform.EnableMenu("772", true);
                                else objform.EnableMenu("772", false);
                            }

                        }
                        else
                        {
                            objform.EnableMenu("772", false);  // Copy
                            objform.EnableMenu("773", false);  // Paste
                        }
                        break;

                    case "PCQTMSTR": //Pre-Cost Quote Master
                        PreCost_QuoteMaster_RightClickEvent(ref eventInfo,ref BubbleEvent);
                        break;
                    case "PCQUOTE": //Pre-Cost Quote Selection  
                        PreCost_QuoteSelection_RightClickEvent(ref eventInfo, ref BubbleEvent);
                        break;
                    case "PCQTETRAN": //Pre-Cost Quote Transaction  
                      //objform.EnableMenu("1283", false);
                        //objform.EnableMenu("1284", false);
                        if (eventInfo.ColUID == "#") objform.EnableMenu("1292", true); //Add Row
                        else objform.EnableMenu("1292", false);
                        break;

                }
            }
            catch (Exception ex)
            {
            }
        }


        private void RightClickMenu_Add(string MainMenu, string NewMenuID, string NewMenuName, int position)
        {
            SAPbouiCOM.Menus omenus;
            SAPbouiCOM.MenuItem omenuitem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;
            oCreationPackage =(SAPbouiCOM.MenuCreationParams)clsModule.objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            omenuitem = clsModule.objaddon.objapplication.Menus.Item(MainMenu); // Data'
            if (!omenuitem.SubMenus.Exists(NewMenuID))
            {
                oCreationPackage.UniqueID = NewMenuID;
                oCreationPackage.String = NewMenuName;
                oCreationPackage.Position = position;
                oCreationPackage.Enabled = true;
                omenus = omenuitem.SubMenus;
                omenus.AddEx(oCreationPackage);
            }
        }

        private void RightClickMenu_Delete(string MainMenu, string NewMenuID)
        {
            SAPbouiCOM.MenuItem omenuitem;
            omenuitem = clsModule.objaddon.objapplication.Menus.Item(MainMenu); // Data'
            if (omenuitem.SubMenus.Exists(NewMenuID))
            {
                clsModule.objaddon.objapplication.Menus.RemoveEx(NewMenuID);
            }
        }

        private void PreCost_QuoteSelection_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                objform =clsModule. objaddon.objapplication.Forms.ActiveForm;
                Matrix0 =(SAPbouiCOM.Matrix) objform.Items.Item("mprodmod").Specific;
                if (eventInfo.BeforeAction)
                {
                    objform.EnableMenu("1283", false); //Remove
                    objform.EnableMenu("1285", false); //Restore                    
                    objform.EnableMenu("1286", false); // Close
                    if (eventInfo.ColUID == "#" ) objform.EnableMenu("1292", true); //Add Row
                    else objform.EnableMenu("1292", false); //Add Row
                    try
                    {
                        // Copy Table                        
                        if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item(eventInfo.ItemUID).Specific;
                            if (eventInfo.Row == 0) objform.EnableMenu("784", true); //Copy Table
                            clsModule.objaddon.objGlobalVariables.contentMatCurRow = eventInfo.Row - 1;
                            if (Matrix0.Columns.Item(eventInfo.ColUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            {
                                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).String != "") objform.EnableMenu("772", true);  // Copy  
                            }
                            else if (Matrix0.Columns.Item(eventInfo.ColUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            {
                                if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).Selected.Value != "") objform.EnableMenu("772", true);  // Copy  
                            }
                            else
                                objform.EnableMenu("772", false);

                        }
                        else if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT || objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                        {
                            if (((SAPbouiCOM.EditText)objform.Items.Item(eventInfo.ItemUID).Specific).String != "") objform.EnableMenu("772", true);  // Copy
                        }
                        else if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            if (((SAPbouiCOM.ComboBox)objform.Items.Item(eventInfo.ItemUID).Specific).Selected.Value != "") objform.EnableMenu("772", true);  // Copy
                        }
                        else
                            if (eventInfo.ItemUID != "") objform.EnableMenu("772", true);
                        else objform.EnableMenu("772", false);
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else
                {     
                    objform.EnableMenu("784", false);
                    objform.EnableMenu("1284", false);
                    objform.EnableMenu("772", false);
                    objform.EnableMenu("1293", false);
                    objform.EnableMenu("1292", false); //Add Row
                }
                
            }
            catch (Exception ex)
            {
            }
        }

        private void PreCost_QuoteMaster_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                //Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mtxcont").Specific;
                
                if (eventInfo.BeforeAction)
                {
                    objform.EnableMenu("1283", false); // Remove
                    objform.EnableMenu("1285", false); // Restore
                    objform.EnableMenu("1284", false); // Cancel
                    objform.EnableMenu("1286", false); // Close
                    if (eventInfo.ColUID == "#" && eventInfo.Row>0)
                    {
                        objform.EnableMenu("1293", true); // Remove Row Menu
                    }
                    objform.EnableMenu("1292", true); //Add Row
                    try
                    {
                        // Copy Table                        
                        if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item(eventInfo.ItemUID).Specific;
                            if (eventInfo.Row == 0) objform.EnableMenu("784", true); //Copy Table
                           clsModule.objaddon.objGlobalVariables.contentMatCurRow = eventInfo.Row-1;
                            if (Matrix0.Columns.Item(eventInfo.ColUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            {
                                if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).String != "") objform.EnableMenu("772", true);  // Copy  
                            }
                            else if (Matrix0.Columns.Item(eventInfo.ColUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            {
                                if (((SAPbouiCOM.ComboBox)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).Selected.Value != "") objform.EnableMenu("772", true);  // Copy  
                            }
                            else
                                objform.EnableMenu("772", false);
                            
                        }
                        else if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EDIT || objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                        {
                            if (((SAPbouiCOM.EditText)objform.Items.Item(eventInfo.ItemUID).Specific).String != "") objform.EnableMenu("772", true);  // Copy
                        }
                        else if (objform.Items.Item(eventInfo.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            if (((SAPbouiCOM.ComboBox)objform.Items.Item(eventInfo.ItemUID).Specific).Selected.Value != "") objform.EnableMenu("772", true);  // Copy
                        }
                        else
                            if (eventInfo.ItemUID!="") objform.EnableMenu("772", true);
                            else objform.EnableMenu("772", false);
                    }
                    catch (Exception ex)
                    {      
                    }
                }
                else
                {
                    objform.EnableMenu("1293", false); // Remove Row Menu
                    objform.EnableMenu("784", false);
                    objform.EnableMenu("1284", false);
                    objform.EnableMenu("772", false);
                    objform.EnableMenu("1293", false);
                }

            }
            catch (Exception ex)
            {
            }
        }

    }
}
