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
    /// Interaction logic for UpdatePartWindow.xaml
    /// </summary>
    public partial class UpdatePartWindow : Window
    {
        PartDBEntities _db = new PartDBEntities();
        int Id;
        public UpdatePartWindow(int partId)
        {
            InitializeComponent();
            Id = partId;
            Part updatePart = (from p in _db.Parts where p.Id == Id select p).Single();
            DescTB.Text = updatePart.Description;
            partNumTB.Text = updatePart.PartNumber;
            linkTB.Text = updatePart.Link;
            supplierTB.Text = updatePart.Supplier;
            priceTB.Text = updatePart.Price;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Part updatePart = (from p in _db.Parts where p.Id == Id select p).Single();
            updatePart.Link = linkTB.Text;
            updatePart.PartNumber = partNumTB.Text;
            updatePart.Description = DescTB.Text;
            updatePart.Supplier = supplierTB.Text;
            updatePart.Price = priceTB.Text;
            _db.SaveChanges();
            MainWindow.dataGrid.ItemsSource = _db.Parts.ToList();
            this.Hide();
        }
    }
}
