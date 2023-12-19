using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Common
{
    class clsModule
    {
        public static clsAddon objaddon;
        public static bool HANA = false;

        [STAThread()]
        public static void Main(string[] args)
        {  
            try
            {
               
                // Application & Company Connection                
                objaddon = new clsAddon();
                objaddon.Intialize(args);


            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error in Module : " + ex.Message.ToString());
              
            }
        }
    }
}
