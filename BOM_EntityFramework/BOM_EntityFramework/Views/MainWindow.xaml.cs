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
        public static Frame frame;
       
        public MainWindow()
        {
            Uri iconUri = new Uri("pack://application:,,,/lightningbolt.ico", UriKind.RelativeOrAbsolute);
            this.Icon = BitmapFrame.Create(iconUri);

            InitializeComponent();
            Main.Content = new PartsHomePage();
            frame = Main;
            //Load();

        }
        private void Load()
        {
            var parts = new PartsViewModel();

            // PartsGrid.ItemsSource = _db.Parts.ToList();
            //dataGrid = PartsGrid;
        }

        private void updateBtn_Click(object sender, RoutedEventArgs e)
        {
            //int Id = (PartsGrid.SelectedItem as Part).Id;
            //UpdatePartWindow Ipage = new UpdatePartWindow(Id);
            //Ipage.ShowDialog();
        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            var delete = MessageBox.Show("Are you sure you want to delete\n this part?","Delete Part",MessageBoxButton.YesNo);

            if (delete.ToString() == "Yes")
            {
                //int Id = (PartsGrid.SelectedItem as Part).Id;
                //Part deletePart = (from p in _db.Parts where p.Id == Id select p).Single();
               
                //_db.Parts.Remove(deletePart);
                //_db.SaveChanges();
                //Load();
            }


        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            //AddPartWindow Ipage = new AddPartWindow();
            //Ipage.ShowDialog();
            Main.Content = new AddPartPage();
            frame = Main;

        }

        private void createBOMBtn_Click(object sender, RoutedEventArgs e)
        {
            Main.Content = new BOMCreationPage();
            frame = Main;

            //BOMCreationWindow Ipage = new BOMCreationWindow();
            //if (!isOpen)
            //{
            //    Ipage.ShowDialog();
            //    //this.Content = Ipage;
            //    isOpen = true;
            //}
        }

        private void SortListByCategoery_Click(object sender, RoutedEventArgs e)
        {

        }

        private void PartsMainBtn_Click(object sender, RoutedEventArgs e)
        {
            Main.Content = new PartsHomePage();
            frame = Main;

        }

        private void AddCatBtn_Click(object sender, RoutedEventArgs e)
        {
            Main.Content = new AddCategoeryPage();
            frame = Main;
        }
    }
}
