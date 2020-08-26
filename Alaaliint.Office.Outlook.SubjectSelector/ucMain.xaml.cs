
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    /// <summary>
    /// Interaction logic for ucMain.xaml
    /// </summary>
    public partial class ucMain : UserControl
    {
        //this.DataContext = this;
        public delegate void MyControlEventHandler(object sender, MyControlEventArgs args);

        public event MyControlEventHandler OnButtonClick;

       

        public delegate void MyButtonResetEventHandler(object sender, Object args);
        public event MyButtonResetEventHandler OnButtonResetClick;

        private SubjectTopic _SelectedSubjectTopic;
        public SubjectTopic SelectedSubjectTopic
        {
            get
            {
                return _SelectedSubjectTopic;
            }
            set
            {
                _SelectedSubjectTopic = value;
                if(_SelectedSubjectTopic != null)
                {
                    filterSubjectType(_SelectedSubjectTopic);
                }
              
            }
        }
        private ObservableCollection<SubjectTopic> _subjectTopicList;
        public ObservableCollection<SubjectTopic> SubjectTopicList
        {
            get
            {
                return _subjectTopicList;
            }
            set
            {
                _subjectTopicList = value;
               
            }
        }
        private ObservableCollection<SubjectType> _subjectTypeFilterList;
        public ObservableCollection<SubjectType> SubjectTypeFilterList
        {
            get
            {
                return _subjectTypeFilterList;
            }
            set
            {
                _subjectTypeFilterList = value;

            }
        }
        private SubjectType _selectedSubjectType;
        public SubjectType SelectedSubjectType
        {
            get
            {
                return _selectedSubjectType;
            }
            set
            {
                _selectedSubjectType = value;

            }
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
        
        public ucMain()
        {
            InitializeComponent();

            this.DataContext = this;

            this.SubjectTypeFilterList = new ObservableCollection<SubjectType>();


        }
       

        private void ButtonClicked(object sender, RoutedEventArgs e)
        {
            
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            MyControlEventArgs retvals = new MyControlEventArgs(true, null, null, null);
            if (sender == btnCancel)
            {
                retvals.IsOK = false;
                retvals.Subject = "";
            }
            else if (sender == btnOK)
            {
                if (string.IsNullOrEmpty(this.Subject) || this.SelectedSubjectTopic == null || this.SelectedSubjectType == null)
                {
                    Mouse.OverrideCursor = null;
                    return;
                }
                    

                retvals.IsOK = true;
                retvals.Subject = this.Subject;//string.Format("[{0:D}-{1:D}-{2:D}] {3:D}", this.SelectedSubjectTopic.Code,this.SelectedSubjectType.Code, "", this.SubjectTextBox.Text); //  //[TopicCode-TypeCode-ID] Subject
                retvals.SubjectTopicValue = this.SelectedSubjectTopic;
                retvals.SubjectTypeValue  = this.SelectedSubjectType;

            }
            if (OnButtonClick != null)
                OnButtonClick(this, retvals);

            Mouse.OverrideCursor = null;
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            this.SubjectTopicList.Clear();

            this.SubjectTypeFilterList.Clear();

            this.Subject = "";

            if (OnButtonResetClick != null)
                OnButtonResetClick(this, null);

            Mouse.OverrideCursor = null;
        }

       
        private void filterSubjectType(SubjectTopic obj)
        {
            this.SubjectTypeFilterList.Clear();
            foreach (SubjectType stype in Globals.ThisAddIn.SubjectTypeList)
            {
                if (stype.Parent.Title == obj.Title)
                {
                    this.SubjectTypeFilterList.Add(stype);
                }
               
            }
        }

        public void reset()
        {
            foreach (SubjectTopic s in this.SubjectTopicList)
            {
                if (s.Code == "000")
                    this.SelectedSubjectTopic = s;
            }
            

            this.Subject = "";
        }
    }

    public class MyControlEventArgs : EventArgs
    {
        private bool _IsOK;
        private string _Subject;
        private SubjectTopic _SubjectTopicValue;
        private SubjectType _SubjectTypeValue;
        public MyControlEventArgs(bool result, string subject,SubjectTopic subjectTopicValue,SubjectType subjectTypeValue)
        {
            _IsOK = result;
            _Subject = subject;
            _SubjectTopicValue = subjectTopicValue;
            _SubjectTypeValue = subjectTypeValue;
        }
        public bool IsOK
        {
            get { return _IsOK; }
            set { _IsOK = value; }
        }
        public string Subject 
        {
            get { return _Subject; }
            set { _Subject = value; }
        }
        public SubjectTopic SubjectTopicValue
        {
            get { return _SubjectTopicValue; }
            set { _SubjectTopicValue = value; }
        }
        public SubjectType SubjectTypeValue
        {
            get { return _SubjectTypeValue; }
            set { _SubjectTypeValue = value; }
        }
    }

  
}
