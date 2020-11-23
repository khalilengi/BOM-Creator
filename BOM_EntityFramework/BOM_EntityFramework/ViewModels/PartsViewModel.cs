using BOM_EntityFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace BOM_EntityFramework.ViewModels
{
    class PartsViewModel  : NotifyPropertyBase
    {
        PartDBEntities context = new PartDBEntities();
    
        Microsoft.Office.Interop.Excel.Application Excel;
        //private ObservableCollection<Part> _partsCollection;
        //public ObservableCollection<Part> PartsCollection
        //{
        //    get { return _partsCollection; }
        //    set
        //    {
        //        _partsCollection = value;
        //        OnPropertyChanged("PartsCollection");
        //    }
        //}
        private List<Part> _partsCollection;
        public List<Part> PartsCollection
        {
            get { return _partsCollection; }
            set
            {
                _partsCollection = value;
               // OnPropertyChanged("PartsCollection");
            }
        }
       
        private List<BOMPart> _BOMPartCollection;
        public List<BOMPart> BOMPartCollection
        {
            get { return _BOMPartCollection; }
            set
            {
                _BOMPartCollection = value;
                // OnPropertyChanged("PartsCollection");
            }
        }
        public void GetParts()
        {
            //var query = from a in context.Parts
            //            select a;

            //var partsList = context.Parts.ToList();
            _partsCollection = context.Parts.ToList();

            //_partsCollection = new ObservableCollection<Part>(partsList);
            //_partsCollection = new ObservableCollection<Part>(query);
        }

        private string _jobNumber;
        public string JobNumber
        {
            get { return _jobNumber; }
            set
            {
                _jobNumber = value;
                OnPropertyChanged("JobNumber");
            }
        }

        private List<ObservableBOMPart> _observableBOMPartsCollection;
        public List<ObservableBOMPart> ObservableBOMPartsCollection
        {
            get { return _observableBOMPartsCollection; }
            set
            {
                _observableBOMPartsCollection = value;
                OnPropertyChanged("ObservableBOMPartsCollection");
            }
        }



        public void GetBOMParts()
        {
            _BOMPartCollection = context.BOMParts.ToList();
        }


        public PartsViewModel()
        {
            GetParts();
        }

        public void AddPartToBOM(Part part)
        {
            if (_observableBOMPartsCollection == null)
            {
                _observableBOMPartsCollection = new List<ObservableBOMPart>();
            }
            ObservableBOMPart addedPart = new ObservableBOMPart(part.Id, part.Description, part.PartNumber, part.Supplier, part.Price);
            try
            {
                _observableBOMPartsCollection.Add(addedPart);
                //BOMGrid.ItemsSource = bomCollection;

            }
            catch (Exception ex)
            {

            }
        }

        public void RemovePartToBOM(ObservableBOMPart part)
        {
            if (_observableBOMPartsCollection == null)
            {
                _observableBOMPartsCollection = new List<ObservableBOMPart>();
            }
            //ObservableBOMPart removedPart = new ObservableBOMPart(part.Id, part.Description, part.PartNumber, part.Supplier, part.Price);
            ObservableBOMPart removedPart = _observableBOMPartsCollection.Where(p => p.PartId == part.PartId).FirstOrDefault();
            try
            {
                _observableBOMPartsCollection.Remove(removedPart);
                //BOMGrid.ItemsSource = bomCollection;

            }
            catch (Exception ex)
            {

            }
        }



        public void ExportBOMExcel()
        {
            GetParts();
            //Setup
            string directoryPath = @"C:\Users\Public\Documents";
            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();
            string BOM = _jobNumber+"_Electrical_BOM_" + date + "_" + time;
            BOM = BOM.Replace('/', '.');
            BOM = BOM.Replace(':', '.');
            string path = directoryPath + $"\\{BOM}_report.xlsx";
            var splitTitle = BOM.Split('_');
            Excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = Excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            Worksheet workSheet = workBook.Worksheets[1];
            int rowNum = 1;
            int colNum = 1;
            int furthestCol = 1;

            BOM = _jobNumber + "_" + splitTitle[1] + "_" + splitTitle[2];
            workSheet.Name = _jobNumber + " BOM";
            //Data To display
            workSheet.Columns.ColumnWidth = 60;

            colNum = 1;

            //Title Bar
            DisplayTitle(workSheet, rowNum, colNum, "Job Number");
            colNum++;
            DisplayData(workSheet, rowNum, colNum, _jobNumber);
            rowNum++;
            colNum = 1;
            DisplayTitle(workSheet, rowNum, colNum, "Date Created");
            colNum++;
            DisplayData(workSheet, rowNum, colNum, date);
            rowNum++;
            colNum=1;
            DisplayTitle(workSheet, rowNum, colNum, "Qty");
            colNum++;
            DisplayTitle(workSheet, rowNum, colNum, "Part Description");
            colNum++;
            DisplayTitle(workSheet, rowNum, colNum, "Part Number");
            colNum++;
            DisplayTitle(workSheet, rowNum, colNum, "Price");
            colNum++;
            DisplayTitle(workSheet, rowNum, colNum, "Supplier");
            colNum++;
            DisplayTitle(workSheet, rowNum, colNum, "Link");
            colNum++;
            
            rowNum++;

            foreach (var item in ObservableBOMPartsCollection)
            {
                colNum = 1;
                //workSheet.Columns[colNum].ColumnWidth = 60;
                //workSheet.Rows[rowNum].RowHeight = 45;
                DisplayData(workSheet, rowNum, colNum, item.Quantity.ToString());
                colNum++;
                DisplayData(workSheet, rowNum, colNum, item.Description);
                colNum++;
                DisplayData(workSheet, rowNum, colNum, item.PartNumber);
                colNum++;
                DisplayData(workSheet, rowNum, colNum, item.Price);
                colNum++;
                DisplayData(workSheet, rowNum, colNum, item.Supplier);
                colNum++;
                string link = _partsCollection.First(p => p.PartNumber == item.PartNumber).Link;
                DisplayData(workSheet, rowNum, colNum, link);
                //colNum++;
             
                furthestCol = colNum;
             
                rowNum++;
            }
            rowNum--;

            // worksheet formatting 
            workSheet.Rows.AutoFit();
            workSheet.Columns.AutoFit();
            //workSheet.Columns.ColumnWidth = 20;

            //string lastCell = "B" + rowNum.ToString();
            //workSheet.get_Range("A2", lastCell).Cells.HorizontalAlignment =
            //        Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            var excelCellrange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowNum, furthestCol]];
            Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            workSheet.PageSetup.RightFooter = "&P/&N";
            workSheet.PageSetup.CenterHeader = _jobNumber+" Electrical BOM";
            workSheet.PageSetup.CenterFooter = "PREFIX CORPORATION, Rochester Hills,MI";
            workSheet.PageSetup.LeftFooter = date;
            workSheet.PageSetup.LeftHeader = "Prefix Corporation";
            workSheet.PageSetup.RightHeader = "Confidential";

            var pageSetup = workSheet.PageSetup;
            pageSetup.Orientation = XlPageOrientation.xlLandscape;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = false; //fitToPageTall has to be false otherwise it will fit to one page 
            pageSetup.Zoom = false; // Zoom will allow to setup print all on one page

            workBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //workBook.Application.Quit();
            Excel.Visible = true;
            Excel.WindowState = XlWindowState.xlMaximized;

            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(Excel);
            Marshal.FinalReleaseComObject(Excel);
        }
        #region Populate Cells Methods
        private static void DisplayTitle(Worksheet workSheet, int rowNum, int colA, string title)
        {
            workSheet.Cells[rowNum, colA] = title;
            workSheet.Cells[rowNum, colA].Font.Bold = true;
            workSheet.Cells[rowNum, colA].Font.Underline = true;
            workSheet.Cells[rowNum, colA].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            var colorCell = workSheet.Cells[rowNum, colA];
            colorCell.Interior.Color = XlRgbColor.rgbLightGray;

        }
        private static void DisplayProject(Worksheet workSheet, int rowNum, string title)
        {
            workSheet.Cells[rowNum, 1] = title;
            workSheet.Cells[rowNum, 1].Font.Bold = true;
            workSheet.Cells[rowNum, 1].Font.Underline = true;
            workSheet.Cells[rowNum, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            var colorCell = workSheet.Cells[rowNum, 1];
            colorCell.Interior.Color = XlRgbColor.rgbLightGray;
            workSheet.Range[workSheet.Cells[rowNum, 1], workSheet.Cells[rowNum, 12]].Merge();

        }
        private static void DisplayData(Worksheet workSheet, int rowNum, int colA, string data)
        {
            if (data == "True")
            {
                //workSheet.Cells[rowNum, colB] = "X";
                workSheet.Cells[rowNum, colA] = "Yes";
            }
            else if (data == "False")
            {
                //workSheet.Cells[rowNum, colB] = "";
                workSheet.Cells[rowNum, colA] = "No";
            }
            else
            {
                workSheet.Cells[rowNum, colA] = data;
            }

            workSheet.Rows[rowNum].WrapText = true;
            var colorCell = workSheet.Cells[rowNum, colA];
            //colorCell.Interior.Color = XlRgbColor.rgbYellow;
            workSheet.Cells[rowNum, colA].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        private static void AddPicture(Worksheet workSheet, int row, int col, string picture)
        {
            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[row, col];
            float Left = (float)((double)oRange.Left) + 5;
            float Top = (float)((double)oRange.Top) + 2;
            const float ImageHeight = 40;
            const float ImageWidth = 40;
            workSheet.Shapes.AddPicture(picture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, ImageWidth, ImageHeight);

        }

        
        #endregion

    }
}
