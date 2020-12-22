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
    /// Interaction logic for AddCategoeryPage.xaml
    /// </summary>
    public partial class AddCategoeryPage : Page
    {
        PartDBEntities _db = new PartDBEntities();
        public AddCategoeryPage()
        {
            InitializeComponent();
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            Catergoery catergoery = new Catergoery()
            {

                Name = NameTB.Text,
                Description = DescTB.Text
            };
            _db.Catergoeries.Add(catergoery);
            _db.SaveChanges();

            MainWindow.frame.Content = new PartsHomePage();
                    
                    
        }

        private void Cancel_Btn_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.frame.Content = new PartsHomePage();
        }
    }
}
