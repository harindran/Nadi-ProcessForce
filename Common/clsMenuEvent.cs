using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Common
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;
        SAPbouiCOM.Form oUDFForm;

        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {

                    switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                    {

                    }
                }
                else
                {
                    switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                    {

                    }
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
                if (pval.BeforeAction == true)
                {
                }

                else
                {

                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
