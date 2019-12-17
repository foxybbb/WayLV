using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
   class LoginHtml 
    {
        public int AuthPR()
        {
            if (webBrowser1.Document == null)
            {
                string message = "Please, Log in";
                string title = "WARNING";
                MessageBox.Show(message, title);
                return 1;
            }
            else
            {
                scrapHtmlTable();
                return 0;
            }


        }


    }
}
