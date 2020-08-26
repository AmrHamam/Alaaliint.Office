using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public partial class frmMain : System.Windows.Forms.Form
    {
        private ElementHost ctrlHost;
        private ucMain wpfCtrl;
        
        public frmMain()
        {
            InitializeComponent();

           
        }
        private string _Subject;
        public string Subject
        {
            get
            {
                return _Subject;
            }
            set
            {
                _Subject = value;
            }
        }
        private Microsoft.Office.Interop.Outlook.MailItem _myMailItem;
        public Microsoft.Office.Interop.Outlook.MailItem myMailItem
        {
            get
            {
                return _myMailItem;
            }
            set
            {
                _myMailItem = value;
            }
        }
        


        private void frmMain_Load(object sender, EventArgs e)
        {
            ctrlHost = new ElementHost();
            ctrlHost.Dock = DockStyle.Fill;
            pnl.Controls.Add(ctrlHost);
            wpfCtrl = new ucMain();
           // wpfCtrl.InitializeComponent();
            ctrlHost.Child = wpfCtrl;
           
            wpfCtrl.OnButtonClick += new ucMain.MyControlEventHandler(Ctrl_OnButtonClick);

            

            wpfCtrl.OnButtonResetClick += new ucMain.MyButtonResetEventHandler(Ctrl_OnButtonResClick);

            this.wpfCtrl.SubjectTopicList = Globals.ThisAddIn.SubjectTopicList;

            this.wpfCtrl.reset();


        }
        private void Ctrl_OnButtonResClick(object sender, Object args)
        {


            Globals.ThisAddIn.Reset();

            this.wpfCtrl.reset();

        }
      

        private void Ctrl_OnButtonClick( object sender,MyControlEventArgs args)
        {
            if (args.IsOK)
            {

                this.Subject = Globals.ThisAddIn.InsertEmailSubjectSerial(args.Subject, "Pending", args.SubjectTopicValue , args.SubjectTypeValue ,myMailItem);
                
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
               
            }
            else
            {
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
        }
    }
}
