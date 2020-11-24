using BOM_EntityFramework.ViewModels;
using BOM_EntityFramework.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace BOM_EntityFramework
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PartDBEntities _db = new PartDBEntities();
        public static DataGrid dataGrid;
        private int _categoeryNum;
        bool isOpen = false;
       
        public MainWindow()
        {
            InitializeComponent();
            Load();

        }
        private void Load()
        {
            var parts = new PartsViewModel();

             PartsGrid.ItemsSource = _db.Parts.ToList();
            dataGrid = PartsGrid;
        }

        private void updateBtn_Click(object sender, RoutedEventArgs e)
        {
            int Id = (PartsGrid.SelectedItem as Part).Id;
            UpdatePartWindow Ipage = new UpdatePartWindow(Id);
            Ipage.ShowDialog();
        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            int Id = (PartsGrid.SelectedItem as Part).Id;
            Part deletePart = (from p in _db.Parts where p.Id == Id select p).Single();
            //var part = PartsGrid.SelectedItem as Part;
            //_db.Parts.Attach(part);
            //_db.Parts.Remove(Part);
            _db.Parts.Remove(deletePart);
            _db.SaveChanges();
            Load();

        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            AddPartWindow Ipage = new AddPartWindow();
            Ipage.ShowDialog();
        }

        private void createBOMBtn_Click(object sender, RoutedEventArgs e)
        {
            
            BOMCreationWindow Ipage = new BOMCreationWindow();
            //BOMCreationPage Ipage = new BOMCreationPage();
            if (!isOpen)
            {
                Ipage.ShowDialog();
                //this.Content = Ipage;
                isOpen = true;
            }
            //Ipage.Show();
        }

        private void SortListByCategoery_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
