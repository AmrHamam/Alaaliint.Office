using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Alaaliint.Office.Outlook.SubjectSelector;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public partial class RibbonSubjectSelector
    {
        private void RibbonSubjectSelector_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            MailItem mailItem = inspector.CurrentItem as MailItem;
            Globals.ThisAddIn.ShowSubjectForm(mailItem);
        }
    }
}
