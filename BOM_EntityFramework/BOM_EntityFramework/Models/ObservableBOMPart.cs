using BOM_EntityFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM_EntityFramework
{
    class ObservableBOMPart : NotifyPropertyBase
    {
        private int _partId;
        public int PartId
        {
            get { return _partId; }
            set
            {
                _partId = value;
                OnPropertyChanged("PartId");
            }
        }

        private int _quantity;
        public int Quantity
        {
            get { return _quantity; }
            set
            {
                _quantity = value;
                OnPropertyChanged("Quantity");
            } 
        }

        private string _description;
        public string Description
        {
            get { return _description; }
            set
            {
                _description = value;
                OnPropertyChanged("Description");
            }
        }
        private string _partNumber;
        public string PartNumber
        {
            get { return _partNumber; }
            set
            {
                _partNumber = value;
                OnPropertyChanged("PartNumber");
            }
          }

        private string _supplier;
        public string Supplier
        {
            get { return _supplier; }
            set
            {
                _supplier = value;
                OnPropertyChanged("Supplier");
            }
        }

        private string _price;
        public string Price 
        {
            get { return _price; }
            set
            {
                _price = value;
                OnPropertyChanged("Price");
            }
        }

        public ObservableBOMPart(int _partId, string _description, string _partNumber, string _supplier, string _price)
        {
            PartId = _partId;
            Description = _description;
            PartNumber = _partNumber;
            Supplier = _supplier;
            Price = _price;
            Quantity = 1; 
        }
    }
}
