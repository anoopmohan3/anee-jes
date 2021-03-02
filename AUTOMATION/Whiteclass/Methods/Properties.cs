using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Whiteclass.Methods
{
    class Properties
    {
    }
  
   
    class pgHomeWindow
    {
        public static String _HomeWindow()
        {
            return (String)Common.FindPageElement("_HomeWindow");
        }
        public static String _form2()
        {
            return (String)Common.FindPageElement("_form2");
        }

        public static String txtname()
        {
            return (String)Common.FindPageElement("pg_HomeWindow.txtname");
        }
        public static String txtpass()
        {
            return (String)Common.FindPageElement("pg_HomeWindow.txtBox_pass");
        }
        public static String btn_Submit()
        {
            return (String)Common.FindPageElement("pg_HomeWindow.btn_Submit");
        }
        public static String btnCancel()
        {
            return (String)Common.FindPageElement("pg_HomeWindow.btnCancel");
        }
        public static String lbl_Output()
        {
            return (String)Common.FindPageElement("pg_HomeWindow.lbl_Output");
        }
        public static String btn_Close()
        {
            return (String)Common.FindPageElement("pg_form2.form2_btn_close");
        }

    }



 
}
