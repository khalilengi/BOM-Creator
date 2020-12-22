using BOM_EntityFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;
using System.IO;

namespace BOM_EntityFramework.ViewModels
{
    class PartsViewModel : NotifyPropertyBase
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

        private List<Catergoery> _categoryCollection;
        public List<Catergoery> CategoryCollection
        {
            get { return _categoryCollection; }
            set
            {
                _categoryCollection = value;
                OnPropertyChanged("CategoryCollection");
            }
        }

        private string _reportTitle;
        public string ReportTitle
        {
            get { return _reportTitle; }
            set
            {
                _reportTitle = value;
                // NotifyPropertyChanged();
                OnPropertyChanged("ReportTitle");
            }
        }
        private string _directoryPath;
        public string DirectoryPath
        {
            get { return _directoryPath; }
            set
            {
                _directoryPath = value;
                //NotifyPropertyChanged();
                OnPropertyChanged("DirectoryPath");
            }
        }
        private string _path;
        public string Path
        {
            get { return _path; }
            set
            {
                _path = value;
                //NotifyPropertyChanged();
                OnPropertyChanged("Path");
            }
        }

        private List<string> _jobNumberCollection;
        public List<string> JobNumberCollection
        {
            get { return _jobNumberCollection; }
            set
            {
                _jobNumberCollection = value;
                OnPropertyChanged("JobNumberCollection");
            }
        }

        private List<string> _catergoryNameCollection;
        public List<string> CatergoryNameCollection
        {
            get { return _catergoryNameCollection; }
            set
            {
                _catergoryNameCollection = value;
                OnPropertyChanged("CatergoryNameCollection");
            }
        }

        private List<string> _partNumberCollection;
        public List<string> PartNumberCollection
        {
            get { return _partNumberCollection; }
            set
            {
                _partNumberCollection = value;
                OnPropertyChanged("PartNumberCollection");
            }
        }

        private bool _exportViewBtnVisibility;
        public bool ExportViewBtnVisibility
        {
            get { return _exportViewBtnVisibility; }
            set
            {
                _exportViewBtnVisibility = value;
                OnPropertyChanged("ExportViewBtnVisibility");
            }
        }
        public PartsViewModel()
        {
            GetParts();
        }

        public void GetParts()
        {
            //var query = from a in context.Parts
            //            select a;

            //var partsList = context.Parts.ToList();
            _partsCollection = context.Parts.ToList();

            _exportViewBtnVisibility = false;
            //_partsCollection = new ObservableCollection<Part>(partsList);
            //_partsCollection = new ObservableCollection<Part>(query);
        }


        public void GetBOMParts()
        {
            _BOMPartCollection = context.BOMParts.ToList();
        }

        public void GetCategoeries()
        {
            _categoryCollection = context.Catergoeries.ToList();
        }

        public void GetCategoryNames()
        {
            GetCategoeries();
            _catergoryNameCollection = new List<string>();
            foreach (var item in CategoryCollection)
            {
                _catergoryNameCollection.Add(item.Name);
            }
        }

        public void SortPartsByCategory(string category)
        {
            GetParts();
            if (category != "" && category != "Show All")
            {
                int id = CategoryCollection.FirstOrDefault(c => c.Name == category).Id;
                PartsCollection = PartsCollection.Where(c => c.CatergoeryId == id).ToList();
            }
        }
        
        public void GetPartNumberCollection()
        {
            GetParts();
            PartNumberCollection = new List<string>();
            foreach (var item in PartsCollection)
            {
                PartNumberCollection.Add(item.PartNumber);
            }
        }

        public bool CheckIfPartNumberCreated(string jobNumber)
        {
            GetPartNumberCollection();
            return PartNumberCollection.Contains(jobNumber);
        }



        public void AddPartToBOM(Part part)
        {
            if (_observableBOMPartsCollection == null)
            {
                _observableBOMPartsCollection = new List<ObservableBOMPart>();
            }
            ObservableBOMPart addedPart = new ObservableBOMPart(part.Id, part.Description, part.PartNumber, part.Supplier, part.Price, part.Link);
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

      public void SaveBOM()
        {
            DateTime dateNow = DateTime.Today;
            if (_observableBOMPartsCollection !=null && _observableBOMPartsCollection.Count > 0)
            {
                foreach (var part in _observableBOMPartsCollection)
                {
                    BOMPart newBOMPart = new BOMPart()
                    {
                        JobNumber = _jobNumber,
                        PartId = part.PartId,
                        Quantity = part.Quantity,
                        DateCreated = dateNow

                    };
                    context.BOMParts.Add(newBOMPart);
                }
                context.SaveChanges();
            }
            else if (_observableBOMPartsCollection.Count < 1)
            {
                MessageBox.Show("There are no parts added to the BOM.\nPlease add parts and try again.");
            }
        }

        public void GetJobNumbers()
        {
            GetBOMParts();
            JobNumberCollection = new List<string>();
            string lastJobNumber = "";
            foreach (var item in BOMPartCollection)
            {
                if (lastJobNumber != item.JobNumber)
                {
                    _jobNumberCollection.Add(item.JobNumber);
                    lastJobNumber = item.JobNumber;
                }  
            }
        }

        public void ViewBOM(string JobNumber)
        {
            ObservableBOMPartsCollection = new List<ObservableBOMPart>();
            GetParts();
            GetBOMParts();
            var jobNumberParts = BOMPartCollection.Where(j => j.JobNumber == JobNumber);
            foreach (var item in jobNumberParts)
            {
                var part = PartsCollection.Where(p => p.Id == item.PartId).FirstOrDefault();
                ObservableBOMPart bomPart = new ObservableBOMPart(item.PartId, part.Description, part.PartNumber, part.Supplier, part.PartNumber, part.Link, item.Quantity);
                //{
                //};
                //bomPart.Quantity = item.Quantity;
                ObservableBOMPartsCollection.Add(bomPart);

            }
            _exportViewBtnVisibility = true;

        }

        public void ExportBOMExcel()
        {
            GetParts();
            //Setup
            string directoryPath = @"C:\Users\Public\Documents";
            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();
            string BOM = _jobNumber + "_Electrical_BOM_" + date + "_" + time;
            BOM = BOM.Replace('/', '.');
            BOM = BOM.Replace(':', '.');
            string path = DirectoryPath + $"\\{BOM}_Report.xlsx";
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
            rowNum+=2;
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
                //string link = _partsCollection.First(p => p.PartNumber == item.PartNumber).Link;
                //DisplayData(workSheet, rowNum, colNum, link);
                DisplayData(workSheet, rowNum, colNum, item.Link);
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
            else if (data.Contains("www.") || data.Contains(".com"))
            {
                workSheet.Hyperlinks.Add(workSheet.Cells[rowNum, colA], data, Type.Missing);
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

        public DirectoryReturn CheckForDirectory(string type)
        {
            string directoryPath = @"C:\Users\Public\Documents";
            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();
            string BOM = _jobNumber + "_Electrical_BOM_" + date + "_" + time;
            BOM = BOM.Replace('/', '.');
            BOM = BOM.Replace(':', '.');
            string path = directoryPath + $"\\{BOM}_Report.xlsx";
            var splitTitle = BOM.Split('_');
            ReportTitle =BOM;

            DirectoryReturn ReturnInfo;
            bool IsAvailable = false;
            bool canContinue = true;
            char[] noNoCharacters = { '*', '[', ']', '?', '|', ':', '<', '>' };
            if (ReportTitle == "")
            {
                string user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                string[] username = user.Split('\\');
                ReportTitle = username[1] + "_" + DateTime.Now.ToShortDateString() + "_" + DateTime.Now.ToLongTimeString();
                ReportTitle = ReportTitle.Replace('/', '.');
                ReportTitle = ReportTitle.Replace(':', '.');
            }
            else
            {
                ReportTitle = ReportTitle.Replace('/', '_');


            }

            if (DirectoryPath == "" || DirectoryPath == null)
            {
                DirectoryPath = @"C:\Users\Public\Documents";
            }

            if (type == "PDF")
            {
                try
                {
                    string[] files = Directory.GetFiles(DirectoryPath, $"{ReportTitle}_REPORT.pdf", SearchOption.TopDirectoryOnly);
                    Path = DirectoryPath + $"\\{ReportTitle}_REPORT.pdf";

                    if (files.Contains(Path))
                    {
                       // Tools.Global.DisplayDialogBox(DialogBox.DialogType.OK, DialogBox.DialogImage.IMPORTANT, "There is a PDF with this name already. \nPlease change the name and export again.");
                        MessageBox.Show("There is a PDF with this name already. \nPlease change the name and export again.");
                        IsAvailable = false;
                    }
                    else
                    {
                        IsAvailable = true;
                    }

                }
                catch
                {
                    IsAvailable = false;
                }

            }

            if (type == "Excel")
            {
                try
                {

                    string[] files = Directory.GetFiles(DirectoryPath, $"{ReportTitle}_REPORT.xlsx", SearchOption.TopDirectoryOnly);
                    Path = DirectoryPath + $"\\{ReportTitle}_REPORT.xlsx";

                    if (files.Contains(Path))
                    {
                        //Tools.Global.DisplayDialogBox(DialogBox.DialogType.OK, DialogBox.DialogImage.IMPORTANT, "There is a excel file with this name already. \nPlease change the name and export again.");
                        MessageBox.Show("There is a excel file with this name already. \nPlease change the name and export again.");
                        IsAvailable = false;
                    }
                    else
                    {
                        IsAvailable = true;
                    }

                }
                catch
                {
                    IsAvailable = false;
                }
            }
            if (type == "Save")
            {
                try
                {

                    string[] files = Directory.GetFiles(DirectoryPath, $"{ReportTitle}_REPORT.xlsx", SearchOption.TopDirectoryOnly);
                    Path = DirectoryPath + $"\\{ReportTitle}_REPORT.xlsx";

                    if (files.Contains(Path))
                    {
                        MessageBox.Show("There is a excel file with this name already. \nPlease change the name and export again.");
                        //Tools.Global.DisplayDialogBox(DialogBox.DialogType.OK, DialogBox.DialogImage.IMPORTANT, "There is a excel file with this name already. \nPlease change the name and export again.");
                        IsAvailable = false;
                        canContinue = false;
                    }
                    else
                    {
                        IsAvailable = true;
                    }

                }
                catch
                {
                    IsAvailable = false;
                    canContinue = false;

                }
                if (canContinue)
                {
                    try
                    {
                        string[] files = Directory.GetFiles(DirectoryPath, $"{ReportTitle}_REPORT.pdf", SearchOption.TopDirectoryOnly);
                        Path = DirectoryPath + $"\\{ReportTitle}_REPORT.pdf";

                        if (files.Contains(Path))
                        {
                            //  Tools.Global.DisplayDialogBox(DialogBox.DialogType.OK, DialogBox.DialogImage.IMPORTANT, "There is a PDF with this name already. \nPlease change the name and export again.");
                            MessageBox.Show("There is a PDF with this name already. \nPlease change the name and export again.");
                            IsAvailable = false;
                        }
                        else
                        {
                            IsAvailable = true;
                        }

                    }
                    catch
                    {
                        IsAvailable = false;
                    }
                }

            }
            foreach (var item in noNoCharacters)
            {
                if (ReportTitle.Contains(item))
                {
                    //Tools.Global.DisplayDialogBox(DialogBox.DialogType.OK, DialogBox.DialogImage.IMPORTANT, "Couldn't save the file. \n-Make sure there is no file with the same name or the file with the same name is open.\n-Make sure the file name does not contain any of the following characters:  <  >  ?  [  ]  :  | or  *");
                    MessageBox.Show("Couldn't save the file. \n-Make sure there is no file with the same name or the file with the same name is open.\n-Make sure the file name does not contain any of the following characters:  <  >  ?  [  ]  :  | or  *");
                    IsAvailable = false;
                    break;
                }
            }
            ReturnInfo = new DirectoryReturn(ReportTitle, Path, IsAvailable);
            return ReturnInfo;
        }

    }
    public class DirectoryReturn
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsAvailable { get; set; }
        public DirectoryReturn(string _fileName, string _filePath, bool _isAvailable)
        {
            FileName = _fileName;
            FilePath = _filePath;
            IsAvailable = _isAvailable;
        }
    }

    public class ReportFields
    {
        public string FieldName { get; set; }
        public bool IsChecked { get; set; }

        public ReportFields(string _fieldName, bool _isChecked)
        {
            FieldName = _fieldName;
            IsChecked = _isChecked;
        }
    }
}
