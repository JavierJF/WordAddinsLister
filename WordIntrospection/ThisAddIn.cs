using System;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace WordIntrospection
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string addins = "";
            int i = 1;
            try
            {
                foreach(Office.COMAddIn addin in this.Application.COMAddIns)
                {
                    addins += "- " + i.ToString() + ": " + addin.Description + "\r\n";
                    i += 1;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("Loaded addins:\r\n" + addins);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
