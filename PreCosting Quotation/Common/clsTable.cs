using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PreCosting_Quotation.Common
{
    class clsTable
    {        
        public void FieldCreation()
        {
            PreCosting_Master();
            PreCostQuote_Selection();
            PreCostQuotation_Transaction();
        }

        #region Master Data Creation

        public void PreCosting_Master()
        {
            #region Fan Model Master

            AddTables("AT_PCMODEL", "Pre-Cost Fan Model Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCMODEL", "Class", "Class", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "I,I", "II,II", "III,III" });
            AddFields("@AT_PCMODEL", "Arrangement", "Arrangement", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "1,1", "2,2", "3,3" });
            AddFields("@AT_PCMODEL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCMODEL", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCMODEL", "Pre-Cost Fan Model Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCMODEL", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Component Master

            AddTables("AT_PCCOMP", "Pre-Cost Component Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCCOMP", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCCOMP", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCCOMP", "Pre-Cost Component Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCCOMP", new[] { "" }, new[] { "Code", "Name" }, false, true, false);


            #endregion

            #region BoughOut Master

            AddTables("AT_PCBOUT", "Pre-Cost BoughOut Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCBOUT", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCBOUT", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCBOUT", "Pre-Cost BoughOut Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCBOUT", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Test Properties Master

            AddTables("AT_PCTPROP", "Pre-Cost TestProp Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCTPROP", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCTPROP", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCTPROP", "Pre-Cost TestProp Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCTPROP", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Packing Charge Master

            AddTables("AT_PCOHEAD", "Pre-Cost Packing Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCOHEAD", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCOHEAD", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCOHEAD", "Pre-Cost Packing Charges", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCOHEAD", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region PreCosting Quote Master

            AddTables("AT_PCQTMSTR", "Pre-Cost Quote Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddTables("AT_PCQTMSTR1", "Pre-Cost Quote Components", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR2", "Pre-Cost Quote Boughtout", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR3", "Pre-Cost Quote TestProp", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR4", "Pre-Cost Quote OverheadProp", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR5", "Pre-Cost Quote Attachments", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR6", "Pre-Cost Quote Spare", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR7", "Pre-Cost Quote Contingency", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR8", "Pre-Cost Quote Warranty", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PCQTMSTR9", "Pre-Cost Quote Negotiation", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);

            AddFields("@AT_PCQTMSTR", "ModelCode", "Model Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR", "Class", "Class", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "I,I", "II,II", "III,III" });
            AddFields("@AT_PCQTMSTR", "Arrangement", "Arrangement", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "1,1", "2,2", "3,3" });
            AddFields("@AT_PCQTMSTR", "MotorKW", "Motor KW", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR", "Poles", "Poles", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCQTMSTR", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);

            AddFields("@AT_PCQTMSTR1", "CompCode", "Component Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR1", "CompName", "Component Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTMSTR1", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity);

            AddFields("@AT_PCQTMSTR2", "BOutCode", "BoughtOut Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR2", "BOutName", "BoughtOut Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@AT_PCQTMSTR3", "TPropCode", "Test Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR3", "TPropName", "Test Property Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@AT_PCQTMSTR4", "OPropCode", "Overhead Prop Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR4", "OPropName", "Overhead Prop Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            //Attachment Table
            AddFields("@AT_PCQTMSTR5", "TrgtPath", "Target Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCQTMSTR5", "SrcPath", "Source Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCQTMSTR5", "Date", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@AT_PCQTMSTR5", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTMSTR5", "FileExt", "File Extension", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTMSTR5", "FreeText", "Free Text", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);


            AddFields("@AT_PCQTMSTR6", "SparCode", "Spare Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR6", "SparName", "Spare Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@AT_PCQTMSTR7", "ContCode", "Contingency Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR7", "ContName", "Contingency Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@AT_PCQTMSTR8", "WarrCode", "Warranty Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR8", "WarrName", "Warranty Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@AT_PCQTMSTR9", "NegoCode", "Negotiation Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTMSTR9", "NegoName", "Negotiation Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddUDO("AT_PCQTMSTR", "Pre-Cost Quote Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCQTMSTR", new[] { "AT_PCQTMSTR1", "AT_PCQTMSTR2", "AT_PCQTMSTR3", "AT_PCQTMSTR4", "AT_PCQTMSTR5", "AT_PCQTMSTR6", "AT_PCQTMSTR7" , "AT_PCQTMSTR8", "AT_PCQTMSTR9" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Spare Master

            AddTables("AT_PCSPAR", "Pre-Cost Spares Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCSPAR", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCSPAR", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCSPAR", "Pre-Cost Spares Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCSPAR", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Contingency Charges Master

            AddTables("AT_PCCCGE", "Pre-Cost Contingency Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCCCGE", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCCCGE", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCCCGE", "Pre-Cost Contingency Charges", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCCCGE", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Warranty Charges Master

            AddTables("AT_PCWCGE", "Pre-Cost Warranty Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCWCGE", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCWCGE", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCWCGE", "Pre-Cost Warranty Charges", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCWCGE", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion

            #region Negotiation Charges Master

            AddTables("AT_PCNCGE", "Pre-Cost Negotiation Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddFields("@AT_PCNCGE", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCNCGE", "Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "Y", true);
            AddUDO("AT_PCNCGE", "Pre-Cost Negotiation Charges", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PCNCGE", new[] { "" }, new[] { "Code", "Name" }, false, true, false);

            #endregion


        }

        #endregion

        #region Document Data Creation

        public void PreCostQuote_Selection()
        {
            AddTables("AT_PCQTESEL", "PreCost Quote Selection Header", SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTables("AT_PCQTESEL1", "PreCost Quote Selection Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            AddFields("@AT_PCQTESEL", "BPCode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@AT_PCQTESEL", "BPName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTESEL", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date);            
            AddFields("@AT_PCQTESEL", "IndustryC", "Industry Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40);
            AddFields("@AT_PCQTESEL", "Reference", "Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);

            AddFields("@AT_PCQTESEL1", "ModelCode", "Model Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTESEL1", "ModelName", "Model Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTESEL1", "Class", "Class", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTESEL1", "Arrangement", "Arrangement", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTESEL1", "Application", "Application", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "A,A", "B,B" });
            AddFields("@AT_PCQTESEL1", "DutyCon", "Duty Condition", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTESEL1", "MotorKW", "Motor KW", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTESEL1", "Poles", "Poles", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTESEL1", "QuoteTran", "Quote Transaction", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTESEL1", "Total", "Transaction Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "CompSum", "Component Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "BoutSum", "BoughtOut Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "TPropSum", "TestProp Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "SChrgSum", "ServCharge Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "OPropSum", "OverheadProp Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "SparSum", "Spares Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "ContSum", "Contingency Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "WarrSum", "Warranty Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTESEL1", "NegoSum", "Negotiation Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);

            AddUDO("AT_PCQTESEL", "PreCost Quote Selection", SAPbobsCOM.BoUDOObjType.boud_Document, "AT_PCQTESEL", new[] { "AT_PCQTESEL1" }, new[] { "DocEntry", "DocNum" }, true,true, true);

        }

        public void PreCostQuotation_Transaction()
        {
            AddTables("AT_PCQTETRAN", "PreCost Qte Tran Header", SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTables("AT_PCQTETRAN1", "PreCost Qte Tran Components", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN2", "PreCost Qte Tran BoughtOut", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN3", "PreCost Qte Tran Test Prop", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN4", "PreCost Qte Tran Serv Chrges", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN5", "PreCost Qte Tran Oerhead Prop", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN6", "PreCost Qte Tran Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN7", "PreCost Qte Tran Spares", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN8", "PreCost Qte Tran Contingency", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN9", "PreCost Qte Tran Warranty", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            AddTables("AT_PCQTETRAN10", "PreCost Qte Tran Negotiation", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
          
            //Header Table
            AddFields("@AT_PCQTETRAN", "BPCode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@AT_PCQTETRAN", "BPName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@AT_PCQTETRAN", "QteSelNo", "Quote Selection No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN", "QteSelEnt", "Quote Selection Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN", "ModelCode", "Model Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN", "ModelName", "Model Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN", "ModelRow", "Model LineID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_PCQTETRAN", "CompSum", "Component Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "BoutSum", "BoughtOut Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "TPropSum", "TestProp Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "SChrgSum", "ServCharge Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "OPropSum", "OverheadProp Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "SparSum", "Spares Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "ContSum", "Contingency Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "WarrSum", "Warranty Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "NegoSum", "Negotiation Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN", "Rounding", "Rounding", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", true);
            AddFields("@AT_PCQTETRAN", "RTotal", "Round After Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);

            //Components Table
            AddFields("@AT_PCQTETRAN1", "CompCode", "Component Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN1", "CompName", "Component Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN1", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity);
            AddFields("@AT_PCQTETRAN1", "Rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@AT_PCQTETRAN1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Bought Out Table
            AddFields("@AT_PCQTETRAN2", "BOutCode", "BoughtOut Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN2", "BOutName", "BoughtOut Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN2", "MarginPer", "Margin Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            AddFields("@AT_PCQTETRAN2", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@AT_PCQTETRAN2", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN2", "MarginTot", "Margin Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN2", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Test Properties Table
            AddFields("@AT_PCQTETRAN3", "TPropCode", "Test Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN3", "TPropName", "Test Property Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN3", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN3", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Service Charges Table

            AddFields("@AT_PCQTETRAN4", "ServName", "Service ChargeName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN4", "Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@AT_PCQTETRAN4", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity);
            AddFields("@AT_PCQTETRAN4", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN4", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Overhead/packing Charges Table
            AddFields("@AT_PCQTETRAN5", "OPropCode", "Overhead Prop Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN5", "OPropName", "Overhead Prop Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN5", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN5", "MarginPer", "Margin Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            AddFields("@AT_PCQTETRAN5", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Attachment Table
            AddFields("@AT_PCQTETRAN6", "TrgtPath", "Target Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCQTETRAN6", "SrcPath", "Source Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PCQTETRAN6", "Date", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@AT_PCQTETRAN6", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTETRAN6", "FileExt", "File Extension", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PCQTETRAN6", "FreeText", "Free Text", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            //Spares
            AddFields("@AT_PCQTETRAN7", "SparCode", "Spare Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN7", "SparName", "Spare Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN7", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN7", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Contingency Charges Table
            AddFields("@AT_PCQTETRAN8", "ContCode", "Contingency Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN8", "ContName", "Contingency Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN8", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN8", "MarginPer", "Margin Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            AddFields("@AT_PCQTETRAN8", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Warranty Charges Table
            AddFields("@AT_PCQTETRAN9", "WarrCode", "Warranty Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN9", "WarrName", "Warranty Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN9", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN9", "MarginPer", "Margin Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            AddFields("@AT_PCQTETRAN9", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);

            //Negotiation Charges Table
            AddFields("@AT_PCQTETRAN10", "NegoCode", "Negotiation Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PCQTETRAN10", "NegoName", "Negotiation Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PCQTETRAN10", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PCQTETRAN10", "MarginPer", "Margin Percent", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            AddFields("@AT_PCQTETRAN10", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);


            AddUDO("AT_PCQTETRAN", "PreCost Quote Transaction", SAPbobsCOM.BoUDOObjType.boud_Document, "AT_PCQTETRAN", new[] { "AT_PCQTETRAN1", "AT_PCQTETRAN2", "AT_PCQTETRAN3", "AT_PCQTETRAN4", "AT_PCQTETRAN5", "AT_PCQTETRAN6", "AT_PCQTETRAN7", "AT_PCQTETRAN8", "AT_PCQTETRAN9", "AT_PCQTETRAN10" }, new[] { "DocEntry", "DocNum" }, true, true, true);

        }

        #endregion

        #region Table Creation Common Functions

        private void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            // var oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception(String.Concat(clsModule.objaddon.objcompany.GetLastErrorDescription() ," ", strDesc));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(String.Concat( ex.Message.ToString() ," ", strDesc));
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedSysObject= 0,string UDTTable="", SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                // If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                // strTab = "@" + strTab
                // End If
                if (!IsColumnExists(strTab, strCol))
                {
                    // If Not oUserFieldMD1 Is Nothing Then
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    // End If
                    // oUserFieldMD1 = Nothing
                    // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;
                    
                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    if (LinkedSysObject != 0)
                        oUserFieldMD1.LinkedSystemObject = LinkedSysObject;// SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulInvoices ;
                    if (UDTTable != "")
                        oUserFieldMD1.LinkedTable = UDTTable;
                    string[] split_char;
                    if (Validvalues !=null)
            {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule. objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short,true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString() + " " + strTab + " " + strCol);
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet=null;
            string strSQL;
            try
            {
                if (clsModule. objaddon.HANA)
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                }
                else
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" + Table + "' AND AliasID = '" + Column + "'";
                }

                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32( oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddKey(string strTab, string strColumn, string strKey, int i)
        {
            var oUserKeysMD = default(SAPbobsCOM.UserKeysMD);

            try
            {
                // // The meta-data object must be initialized with a
                // // regular UserKeys object
                oUserKeysMD =(SAPbobsCOM.UserKeysMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

                if (!oUserKeysMD.GetByKey("@" + strTab, i))
                {

                    // // Set the table name and the key name
                    oUserKeysMD.TableName = strTab;
                    oUserKeysMD.KeyName = strKey;

                    // // Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn;
                    oUserKeysMD.Elements.Add();
                    oUserKeysMD.Elements.ColumnAlias = "RentFac";

                    // // Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

                    // // Add the key
                    if (oUserKeysMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool Cancel = false, bool canlog = false, bool Manageseries = false)
        {

           SAPbobsCOM.UserObjectsMD oUserObjectMD=null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
               
                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    if(Cancel)
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);
                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
            {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }

                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        #endregion

    }
}
