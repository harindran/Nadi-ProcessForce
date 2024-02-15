using Common.Common;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessForce.Bussiness_Logic
{
    class SalesOrder
    {        
        public void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Form oForm = clsModule.objaddon.objapplication.Forms.Item(oFormUID);

                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                           // Create_Customize_Fields(oForm);
                            break;                     
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }



        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm = clsModule.objaddon.objapplication.Forms.Item(pVal.FormUID);

                if (!pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            // Create_Customize_Fields(oForm);
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            switch (pVal.FormTypeEx)
                            {
                                case "139":
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        UpdateRevisionLogic(oForm);
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                        clsModule.objaddon.objapplication.Menus.Item("1304").Activate();
                                        clsModule.objaddon.objapplication.StatusBar.SetText("Operation Compeletly successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                                    }
                                    break;
                            }
                            break;

                    }
                }              
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        private void Create_Customize_Fields(SAPbouiCOM.Form oForm)
        {
            try
            {
                switch (oForm.TypeEx)
                {
                    case "139":
                        break;
                    default:
                        return;
                }

                SAPbouiCOM.Item oItem;
                clsModule.objaddon.objglobalmethods.WriteErrorLog("Customize Field Start");
                try
                {
                    if (oForm.Items.Item("btnRev").UniqueID == "btnRev")
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {

                }
                switch (oForm.TypeEx)
                {
                    case "139":
                        oItem = oForm.Items.Add("btnRev", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        SAPbouiCOM.Button button = (SAPbouiCOM.Button)oItem.Specific;
                        button.Caption = "Update Revision";
                        oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5;
                        oItem.Top = oForm.Items.Item("2").Top;
                        oItem.Height = oForm.Items.Item("2").Height;
                        oItem.LinkTo = "2";
                        Size Fieldsize = System.Windows.Forms.TextRenderer.MeasureText("Update Revision", new Font("Arial", 12.0f));
                        oItem.Width = Fieldsize.Width;
                        oForm.Items.Item("btnRev").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                        oForm.Items.Item("btnRev").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                        break;
                    default:
                        return;
                }
                oForm.Freeze(false);
                clsModule.objaddon.objglobalmethods.WriteErrorLog("Customize Field Completed");
            }
            catch (Exception ex)
            {
            }
        }


        public bool UpdateRevisionLogic(SAPbouiCOM.Form oForm)
        {
            try
            {

           
            int Docentry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0));

            GeneralService oGeneralService;
            GeneralData oGeneralData;
            GeneralDataParams oGeneralParams;
            GeneralDataCollection oGeneralDataCollection;
            GeneralData oChild;

            oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("CT_PF_ItemDetails");
            oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
            oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralDataCollection = oGeneralData.Child("CT_PF_IDT1");
            //oChild = oGeneralDataCollection.Add();
            //
            SAPbobsCOM.Documents objdocument = null;
            objdocument = (SAPbobsCOM.Documents)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            if (objdocument.GetByKey(Docentry))
            {
                for(int j=0;j< objdocument.Lines.Count;j++)
                {                   
                    objdocument.Lines.SetCurrentLine(j);
                    string itemcode=objdocument.Lines.ItemCode;
                    string Revisioncode;
                    string SeriesName;
                    string lstrquery = "Select \"Code\" from \"@CT_PF_OIDT\" where \"U_ItemCode\"='" + itemcode + "'";
                    
                    bool isupdate = false;
                    Revisioncode = clsModule.objaddon.objglobalmethods.getSingleValue(lstrquery);

                       
                        try
                    {
                        oGeneralParams.SetProperty("Code", Revisioncode);
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                        oGeneralData.SetProperty("Code", Revisioncode);
                        isupdate = true;
                    }
                    catch (Exception ex)
                    {
                        isupdate = false;
                    }

                    lstrquery = "SELECT \"SeriesName\"  FROM nnm1 WHERE \"Series\" ='" + objdocument.Series + "';";
                   SeriesName = clsModule.objaddon.objglobalmethods.getSingleValue(lstrquery);

                    string Rcode = "";
                    Rcode = SeriesName + "-" + objdocument.DocNum + "-" + (j+1).ToString();

                        if (Rcode != objdocument.Lines.UserFields.Fields.Item("U_Revision").Value.ToString())
                        {
                            oGeneralData.SetProperty("U_ItemCode", itemcode);

                            clsModule.objaddon.objglobalmethods.WriteErrorLog(Rcode);

                            int Rc = oGeneralData.Child("CT_PF_IDT1").Count;
                            oGeneralData.Child("CT_PF_IDT1").Add();
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_Code", Rcode);
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_Code", Rcode);
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_LineNum", Rc + 1);
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_ParentItemCode", itemcode);
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_Description", Rcode);
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_Status", "ACT");
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_IsMrpDefault", "Y");
                            oGeneralData.Child("CT_PF_IDT1").Item(Rc).SetProperty("U_Default", "Y");

                            if (isupdate)
                            {
                                oGeneralService.Update(oGeneralData);
                            }
                            else
                            {
                                oGeneralParams = oGeneralService.Add(oGeneralData);
                            }

                            objdocument.Lines.UserFields.Fields.Item("U_Revision").Value = Rcode;
                            objdocument.Lines.UserFields.Fields.Item("U_RevisionName").Value = Rcode;
                        }
                        
                }
                    objdocument.Update();
            }
              
            }
            catch (Exception ex)
            {

                
            }
            return true;
        }





    }

}
