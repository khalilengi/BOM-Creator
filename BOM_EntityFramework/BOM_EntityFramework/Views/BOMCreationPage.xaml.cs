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
    /// Interaction logic for BOMCreationPage.xaml
    /// </summary>
    public partial class BOMCreationPage : Page
    {
        PartDBEntities _db = new PartDBEntities();
        DataGrid dataGrid = new DataGrid();
        DataGrid dataGrid2 = new DataGrid();
        PartsViewModel parts;
        private string lastCategory = "";

        public BOMCreationPage()
        {
            InitializeComponent();
            Load();
        }
      
        private void Load()
        {
            parts = new PartsViewModel();
            //var parts = new PartsViewModel();

            PartsGrid.ItemsSource = _db.Parts.ToList();

            parts.GetCategoryNames();
            parts.CatergoryNameCollection.Add("Show All");
            SortCategoryComboBox.ItemsSource = parts.CatergoryNameCollection;
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            var part = PartsGrid.SelectedItem as Part;
            var bomItemSource = BOMGrid.ItemsSource;
            //List<ObservableBOMPart> bomCollection = new List<ObservableBOMPart>();;
            List<ObservableBOMPart> bomCollection = (List<ObservableBOMPart>)bomItemSource;
            parts.AddPartToBOM(part);

            BOMGrid.ItemsSource = null;
            BOMGrid.ItemsSource = parts.ObservableBOMPartsCollection;
            #region Delete
            //ObservableBOMPart addedPart = new ObservableBOMPart(part.Id, part.Description, part.PartNumber, part.Supplier, part.Price);
            //try
            //{
            //    bomCollection.Add(addedPart);
            //    BOMGrid.ItemsSource = bomCollection;

            //}
            //catch ( Exception ex)
            //{
            //    List<ObservableBOMPart> bomCollection2 = new List<ObservableBOMPart>();
            //    if (bomCollection != null)
            //    {

            //        bomCollection2 = bomCollection;
            //    }
            //    bomCollection2.Add(addedPart);
            //    BOMGrid.ItemsSource = bomCollection2;
            //}
            ////var test = BOMGrid.ItemsSource;
            //dataGrid2 = new DataGrid();
            //dataGrid2 = BOMGrid;
            #endregion

        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            var bomItemSource = BOMGrid.SelectedItem as ObservableBOMPart;
            //List<ObservableBOMPart> bomCollection = new List<ObservableBOMPart>();;
            // List<ObservableBOMPart> bomCollection = (List<ObservableBOMPart>)bomItemSource; ;
            parts.RemovePartToBOM(bomItemSource);
            BOMGrid.ItemsSource = null;
            BOMGrid.ItemsSource = parts.ObservableBOMPartsCollection;
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            parts.JobNumber = JobNumberTB.Text as string;
            if (parts.JobNumber != null && parts.JobNumber != "")
            {
                DirectoryReturn directoryReturn = parts.CheckForDirectory("Excel");
                if (directoryReturn.IsAvailable)
                {
                    parts.ExportBOMExcel();
                }
            }
            
            else
            {
                MessageBox.Show($"Please enter a job number in Job Number textbox\nand try to save and export again");

            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {

           // this.Close();
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            
            parts.JobNumber = JobNumberTB.Text as string;
            if (parts.JobNumber != null && parts.JobNumber != "")
            {
                parts.SaveBOM();
                MessageBox.Show($"{parts.JobNumber} BOM is save to database");
            }
            else
            {
                MessageBox.Show($"Please enter a job number in Job Number textbox\nand try to save again");
                    
            }
        }

        private void SaveExportBtn_Click(object sender, RoutedEventArgs e)
        {
            parts.JobNumber = JobNumberTB.Text as string;

            if (parts.JobNumber != null && parts.JobNumber != "")
            {
                DirectoryReturn directoryReturn = parts.CheckForDirectory("Excel");
                if (directoryReturn.IsAvailable)
                {
                    parts.SaveBOM();
                    parts.ExportBOMExcel();
                }
            }
            else
            {
                MessageBox.Show($"Please enter a job number in Job Number textbox\nand try to save and export again");

            }
        }

        private void SortListByCategory_Click(object sender, RoutedEventArgs e)
        {
            string selectedCategory = SortCategoryComboBox.SelectedItem as string;
            if (selectedCategory != null)
            {
                if (selectedCategory != lastCategory)
                {

                    parts.SortPartsByCategory(selectedCategory);
                    //parts.GetParts();
                    //if (selectedCategory != "" && selectedCategory != "Show All")
                    //{
                    //    int id = parts.CategoryCollection.FirstOrDefault(c => c.Name == selectedCategory).Id;
                    //    parts.PartsCollection = parts.PartsCollection.Where(c => c.CatergoeryId == id).ToList();

                    //}
                }
                PartsGrid.ItemsSource = parts.PartsCollection;
                //dataGrid = PartsGrid;
            }
        }
    }
}
