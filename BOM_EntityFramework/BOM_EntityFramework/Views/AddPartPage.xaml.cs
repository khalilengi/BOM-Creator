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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BOM_EntityFramework.Views
{
    /// <summary>
    /// Interaction logic for AddPartPage.xaml
    /// </summary>
    public partial class AddPartPage : Page
    {
        PartDBEntities _db = new PartDBEntities();

        public AddPartPage()
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

            descTextBox.Text = "";
            partNumTextBox.Text = "";
            linkTextBox.Text = "";
            priceTextBox.Text = "";
            supplierTextBox.Text = "";
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            PartsViewModel parts = new PartsViewModel();
            int i = 9;
            try
            {
                if ((string)catergoeryComboBox.SelectedItem != "")
                {
                    var categoery = _db.Catergoeries.ToList();
                    i = categoery.First(c => c.Name == (string)catergoeryComboBox.SelectedItem).Id;
                }
            }
            catch (Exception ex)
            {

                //throw;
            }
            if (descTextBox.Text == "")
            {
                MessageBox.Show("Please enter a description to this part before adding");
            }
            else if (partNumTextBox.Text == "")
            {
                MessageBox.Show("Please enter a part number to this part before adding");

            }
            else
            {
                bool exists = parts.CheckIfPartNumberCreated(partNumTextBox.Text);
                if (!exists)
                {

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
                    MainWindow.frame.Content = new PartsHomePage();
                    MessageBox.Show($"Part Number {newPart.PartNumber} was added to the database");
                }

                else
                {
                    MessageBox.Show("There is already a part created with this part number.\nCheck parts list to update this part\nor change part number");
                }
                //Load();

            }
        }
    }
}
