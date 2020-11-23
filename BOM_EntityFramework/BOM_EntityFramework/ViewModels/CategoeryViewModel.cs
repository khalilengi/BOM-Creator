using BOM_EntityFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM_EntityFramework.ViewModels
{
    class CategoeryViewModel  : NotifyPropertyBase
    {
        PartDBEntities context = new PartDBEntities();
        private List<Catergoery> _catergoeryCollection;
        public List<Catergoery> CatergoeryCollection
        {
            get { return _catergoeryCollection; }
            set
            {
                _catergoeryCollection = value;
                OnPropertyChanged("PartsCollection");
            }
        }

        public void GetCatergoeries()
        {
            _catergoeryCollection = context.Catergoeries.ToList();
        }

        public CategoeryViewModel()
        {
            GetCatergoeries();
        }


    }
}
