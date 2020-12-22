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
    /// Interaction logic for UpdatePartPage.xaml
    /// </summary>
    public partial class UpdatePartPage : Page
    {
        PartDBEntities _db = new PartDBEntities();
        int Id;
        public UpdatePartPage(int partId)
        {
            InitializeComponent();
            Load(partId);
        }

        private void Load(int partId)
        {
            Id = partId;
            Part updatePart = (from p in _db.Parts where p.Id == Id select p).Single();
            DescTB.Text = updatePart.Description;
            partNumTB.Text = updatePart.PartNumber;
            linkTB.Text = updatePart.Link;
            supplierTB.Text = updatePart.Supplier;
            priceTB.Text = updatePart.Price;

            var categoery = _db.Catergoeries.ToList();
            categoery.Add(new Catergoery()
            {
                Description = ""
            });
            List<string> categoeryNames = categoery.Select(c => c.Name).ToList();
            categoeryCB.ItemsSource = categoeryNames;
          //  categoeryCB.SelectedItem = categoeryNames.Where(c => c == "");
            categoeryCB.SelectedItem = categoeryNames[(int)updatePart.CatergoeryId-1];

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            int i = 1;
            if ((string)categoeryCB.SelectedItem != "" && (string)categoeryCB.SelectedItem != null)
            {
                var categoery = _db.Catergoeries.ToList();
                i = categoery.First(c => c.Name == (string)categoeryCB.SelectedItem).Id;
            }
            Part updatePart = (from p in _db.Parts where p.Id == Id select p).Single();
            updatePart.Link = linkTB.Text;
            updatePart.PartNumber = partNumTB.Text;
            updatePart.Description = DescTB.Text;
            updatePart.Supplier = supplierTB.Text;
            updatePart.Price = priceTB.Text;

            updatePart.CatergoeryId = i;

            _db.SaveChanges();
            MainWindow.frame.Content = new PartsHomePage();
       
        }

        private void Cancel_Btn_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.frame.Content = new PartsHomePage();
        }
    }
}
