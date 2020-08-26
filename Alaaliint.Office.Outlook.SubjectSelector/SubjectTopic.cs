using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public class SubjectTopic : INotifyPropertyChanged
    {
        


        public event PropertyChangedEventHandler PropertyChanged;
        
      
        public SubjectTopic()
        {
        }

        public SubjectTopic(int id ,string Code, string Title)
        {
            this.ID = id;
            this.Code = Code;
            this.Title = Title;
           
        }
        private int _ID;
        public int ID
        {
            get { return _ID; }
            set
            {
                _ID = value;
                // Call OnPropertyChanged whenever the property is updated
                OnPropertyChanged();
            }
        }
        private string _Code;
        public string Code
        {
            get { return _Code; }
            set
            {
                _Code = value;
                // Call OnPropertyChanged whenever the property is updated
                OnPropertyChanged();
            }
        }
        private string _Title;
        public string Title
        {
            get { return _Title; }
            set
            {
                _Title = value;
                // Call OnPropertyChanged whenever the property is updated
                OnPropertyChanged();
            }
        }
      
        // Create the OnPropertyChanged method to raise the event
        // The calling member's name will be used as the parameter.
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
