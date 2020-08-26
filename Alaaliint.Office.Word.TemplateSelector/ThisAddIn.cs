using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using OfficeDevPnP.Core;

namespace Alaaliint.Office.Word.TemplateSelector
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
        private Microsoft.Office.Interop.Word.Table GiveMeATable(
    int columnCount, int rowCount, Microsoft.Office.Interop.Word.Document doc, Microsoft.Office.Interop.Word.Range rg)
        {
            if (columnCount > 0 & rowCount > 0)
            {
                //We aren't goofing around. Let's make a table.
                Microsoft.Office.Interop.Word.Table tbl = default(Microsoft.Office.Interop.Word.Table);
                tbl = doc.Tables.Add(rg, rowCount, columnCount,
                     Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                     Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent);
                return tbl;
            }
            return null;
        }

        private void InsertFromSharePoint()
        {

            //Create a table
            Microsoft.Office.Interop.Word.Table tbl = GiveMeATable(1, 1,
                 this.Application.ActiveDocument, this.Application.Selection.Range);

            //Add some table stylings
            tbl.ApplyStyleHeadingRows = true;
            tbl.ApplyStyleRowBands = true;
            tbl.set_Style("Grid Table 4 - Accent 6");
            //Green. I like green.



            string siteUrl = Properties.Settings.Default.ClientContextUrl;
            using (var context = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
            {

                context.Load(context.Web);
                context.ExecuteQuery();

                //We need a counter...
                int i = 1;
                //And a table header
                tbl.Cell(i, 1).Range.Text = "SharePoint Site Lists - " +
                    context.Web.Title;

                //Requery to get the lists
                context.Load(context.Web.Lists);
                context.ExecuteQuery();
                
            }

        }

    }
}
