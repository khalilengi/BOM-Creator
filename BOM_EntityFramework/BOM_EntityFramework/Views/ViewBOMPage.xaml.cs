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
    /// Interaction logic for ViewBOMPage.xaml
    /// </summary>
    public partial class ViewBOMPage : Page
    {
        PartsViewModel parts;
        public ViewBOMPage()
        {
            InitializeComponent();
            LoadJobNumberCollection();
        }

        public void LoadJobNumberCollection()
        {
            parts = new PartsViewModel();
            parts.GetJobNumbers();
            JobNumberListBox.ItemsSource = null;
            JobNumberListBox.ItemsSource = parts.JobNumberCollection;
        }

        private void ViewBOMBtn_Click(object sender, RoutedEventArgs e)
        {
            string jobNum = JobNumberListBox.SelectedItem as string;
            parts.JobNumber = jobNum;
            parts.ViewBOM(jobNum);
            BOMPartsGrid.ItemsSource = null;
            BOMPartsGrid.ItemsSource = parts.ObservableBOMPartsCollection;

        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
                //TODO - messagebox to export bom 
                
                DirectoryReturn directoryReturn = parts.CheckForDirectory("Excel");
                if (directoryReturn.IsAvailable)
                {
                    parts.ExportBOMExcel();
                }
            
        }
    }
}
