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
    /// Interaction logic for PartsHomePage.xaml
    /// </summary>
    public partial class PartsHomePage : Page
    {
        PartDBEntities _db = new PartDBEntities();
        public static DataGrid dataGrid;
        private int _categoeryNum;
        bool isOpen = false;
        PartsViewModel parts;
        private string lastCategory = "";

        public PartsHomePage()
        {
            InitializeComponent();
            Load();
        }
        
        private void Load()
        {
            parts = new PartsViewModel();
            parts.GetCategoryNames();
            parts.CatergoryNameCollection.Add("Show All");
            SortCategoryComboBox.ItemsSource = parts.CatergoryNameCollection;
            PartsGrid.ItemsSource = _db.Parts.ToList();
            dataGrid = PartsGrid;
            
        }

        private void updateBtn_Click(object sender, RoutedEventArgs e)
        {
            int Id = (PartsGrid.SelectedItem as Part).Id;
            //UpdatePartWindow Ipage = new UpdatePartWindow(Id);
            //Ipage.ShowDialog();
            MainWindow.frame.Content = new UpdatePartPage(Id);
        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            var delete = MessageBox.Show("Are you sure you want to delete\n this part?", "Delete Part", MessageBoxButton.YesNo);

            if (delete.ToString() == "Yes")
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
            string selectedCategory = SortCategoryComboBox.SelectedItem as string;
            if (selectedCategory != null)
            {
                if (selectedCategory != lastCategory)
                {


                    parts.GetParts();
                    if (selectedCategory != "" && selectedCategory != "Show All")
                    {
                        int id = parts.CategoryCollection.FirstOrDefault(c => c.Name == selectedCategory).Id;
                        parts.PartsCollection = parts.PartsCollection.Where(c => c.CatergoeryId == id).ToList();

                    }
                }
                PartsGrid.ItemsSource = parts.PartsCollection;
                dataGrid = PartsGrid;
            }
        }
    }
}
