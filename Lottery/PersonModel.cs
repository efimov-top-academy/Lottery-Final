using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Lottery
{
    public class PersonModel : INotifyPropertyChanged
    {
        Person currentPerson;

        public Person CurrentPerson
        {
            get => currentPerson;
            set
            {
                currentPerson = value;
                OnPropertyChanged(nameof(CurrentPerson));
            }
        }
        public ObservableCollection<Person> Persons { get; set; }
        public PersonModel(IEnumerable<Person> persons)
        {
            Persons = new ObservableCollection<Person>(persons);
            currentPerson = Persons.FirstOrDefault();
        }



        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
