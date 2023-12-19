using Common.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessForce.Bussiness_Logic
{
    class PurchaseRequest
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
                            LoadData(oForm);
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

        public void LoadData(SAPbouiCOM.Form oForm)
        {
            switch (oForm.TypeEx)
            {
                case "1470000200":
                    break;
                default:
                    return;
            }
            SAPbouiCOM.Matrix mt = ((SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific);
            
            for (int i = 1; i < mt.RowCount; i++)
            {
                string ss =((SAPbouiCOM.EditText)(mt.Columns.Item("U_DocEntry").Cells.Item(i).Specific)).Value;

                string lstr = " SELECT cpo.\"U_ItemCode\" , cpo.\"U_Revision\" , cp1.\"U_DocNo\" ,cp1.\"U_LineNum\",cpo.\"U_Revision\",cpo.\"U_RevisionName\"  FROM \"@CT_PF_OMOR\" cpo ";
                lstr += " INNER JOIN \"@CT_PF_MOR6\" cp1 ON cpo.\"DocEntry\" = cp1.\"DocEntry\" ";
                lstr += "WHERE cpo.\"DocEntry\" = '" + ss + "'; ";
                
                clsModule.objaddon.objglobalmethods.WriteErrorLog(ss);
                clsModule.objaddon.objglobalmethods.WriteErrorLog(lstr);

                SAPbobsCOM.Recordset rs =clsModule.objaddon.objglobalmethods.GetmultipleValueRS(lstr);

                if (rs.RecordCount>0)
                {
                  
                    columnEdit(true,mt);
                    ((SAPbouiCOM.EditText)(mt.Columns.Item("U_OrderNo").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_DocNo").Value.ToString();                    
                    ((SAPbouiCOM.EditText)(mt.Columns.Item("U_SubItemCode").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_ItemCode").Value.ToString();
                    ((SAPbouiCOM.EditText)(mt.Columns.Item("U_Revision").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_Revision").Value.ToString();
                    ((SAPbouiCOM.EditText)(mt.Columns.Item("U_OrderLineNo").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_LineNum").Value.ToString();
                     ((SAPbouiCOM.EditText)(mt.Columns.Item("U_Revision").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_Revision").Value.ToString();
                     ((SAPbouiCOM.EditText)(mt.Columns.Item("U_RevisionName").Cells.Item(i).Specific)).Value = rs.Fields.Item("U_RevisionName").Value.ToString();

                    oForm.Items.Item("1470002179").Click();
                    columnEdit(false,mt);

                }
            }
          
        }


        private void columnEdit(bool pblnenable,SAPbouiCOM.Matrix mt)
        {
            mt.Columns.Item("U_OrderNo").Editable = pblnenable;
            mt.Columns.Item("U_SubItemCode").Editable = pblnenable;
            mt.Columns.Item("U_Revision").Editable = pblnenable;
            mt.Columns.Item("U_OrderLineNo").Editable = pblnenable;
        }
    }
}
