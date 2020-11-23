using BOM_EntityFramework.ViewModels;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace BOM_EntityFramework.Views
{
    /// <summary>
    /// Interaction logic for AddPartWindow.xaml
    /// </summary>
    public partial class AddPartWindow : Window
    {
        PartDBEntities _db = new PartDBEntities();

        public AddPartWindow()
        {
            InitializeComponent();
            Load();

        }
       private void Load()
        {
            var categoery = _db.Catergoeries.ToList();
            categoery.Add(new Catergoery()
            {
                Description = ""
            });
            List<string> categoeryNames = categoery.Select(c => c.Name).ToList();
            catergoeryComboBox.ItemsSource = categoeryNames;
            catergoeryComboBox.SelectedItem = categoeryNames.Where(c => c == "");
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            int i = 1;
            //if((string)catergoeryComboBox.SelectedItem != "")
            //{
            //    var categoery = _db.Catergoeries.ToList();
            //    i = categoery.First(c => c.Name == (string)catergoeryComboBox.SelectedItem).Id;
            //}
            Part newPart = new Part()
            {
                Description = descTextBox.Text,
                PartNumber = partNumTextBox.Text,
                Link = linkTextBox.Text,
                Supplier = supplierTextBox.Text,
                Price = priceTextBox.Text,
                CatergoeryId = i
             
            };

            _db.Parts.Add(newPart);
            _db.SaveChanges();
            MainWindow.dataGrid.ItemsSource = _db.Parts.ToList();
            this.Hide();

        }
    }
}
