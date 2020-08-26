using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public class SubjectType : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public SubjectType()
        {
        }

        public SubjectType(string Code, string Title,SubjectTopic Parent)
        {
            this.Code = Code;
            this.Title = Title;
            this.Parent = Parent;
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
        private SubjectTopic _Parent;
        public SubjectTopic Parent
        {
            get { return _Parent; }
            set
            {
                _Parent = value;
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
