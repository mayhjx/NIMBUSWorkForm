using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string CustomFontFamily = "Times New Roman";
        private const double ColumnTitleHeight = 25d; // 96孔板列标题行高
        private const double RowTitleWidth = 25d; // 96孔板行标题列宽
        private const double CellRowHeight = 65d; // 96孔板单元格行高
        private const double CellColWidth = 120d;  // 96孔板单元格列宽
        private const double RowHeight = 30d; // 工作清单header和footer行高

        private List<Sample> Samples = new List<Sample>();
        private List<string> Plates = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            this.SizeToContent = SizeToContent.WidthAndHeight;
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开NIMBUS移液平台工作清单
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "打开NIMBUS移液平台工作清单";
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";

            if (openFileDialog.ShowDialog().GetValueOrDefault() != true)
            {
                return;
            }

            FileName.Text = openFileDialog.SafeFileName;

            // 打开每日操作清单
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "打开每日操作清单";
            openFileDialog1.DefaultExt = ".xls";
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog().GetValueOrDefault() != true)
            {
                return;
            }

            FileName1.Text = openFileDialog1.SafeFileName;

            // 用户再次点击上传文件时清空之前的数据
            Samples = new List<Sample>();
            Plates = new List<string>();
            DocumentPage.Document = null;

            try
            {
                // 读取数据
                Samples = ReadNIMBUSWorkBook(openFileDialog.FileName);
                Samples = ReadOperateWorkBook(Samples, openFileDialog1.FileName);

                // 生成96孔板工作清单
                var flowDocument = new FlowDocument()
                {
                    ColumnWidth = RowTitleWidth + CellColWidth * 12 + 50,
                };
                foreach (var plate in Plates)
                {
                    flowDocument.Blocks.Add(new BlockUIContainer(Create96PlateForm(plate)) { });
                }
                DocumentPage.Document = flowDocument;
            }
            catch (IOException)
            {
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成96孔板工作清单时出现错误：{ex.Message}");
                return;
            }

            PrintButton.IsEnabled = true;
            OutputButton.IsEnabled = true;

            DocumentPage.Width = RowTitleWidth + CellColWidth * 12 + 50;
            DocumentPage.Height = ColumnTitleHeight + CellRowHeight * 8 + RowHeight * 3 + 70;
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog dialog = new PrintDialog();
            dialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
            dialog.PrintTicket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);

            if (dialog.ShowDialog().GetValueOrDefault() == true)
            {
                dialog.PrintDocument(((IDocumentPaginatorSource)DocumentPage.Document).DocumentPaginator, "");
            }
        }

        private void OutputButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog().GetValueOrDefault() == true)
            {
                string filePath = saveFileDialog.FileName;
                CreateBatchListAndWarmInfoWorkBook(filePath);
            }
        }

        private void CreateBatchListAndWarmInfoWorkBook(string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet;
            int rowIndex = 0;

            // 插入上机列表
            foreach (var plate in Plates)
            {
                sheet = workbook.CreateSheet(plate);

            }

            // 插入警告信息
            sheet = workbook.CreateSheet("警告信息");
            rowIndex = 0;
            var row = sheet.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("实验号");
            row.CreateCell(1).SetCellValue("板号");
            row.CreateCell(2).SetCellValue("孔位");
            row.CreateCell(3).SetCellValue("条码");
            row.CreateCell(4).SetCellValue("警告信息");

            foreach (var sample in Samples)
            {
                if (sample.WarmInfo == "1" || sample.WarmInfo == "16384")
                {
                    row = sheet.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(sample.Number);
                    row.CreateCell(1).SetCellValue(sample.Plate);
                    row.CreateCell(2).SetCellValue(sample.Position);
                    row.CreateCell(3).SetCellValue(sample.BarCode);
                    row.CreateCell(4).SetCellValue(sample.WarmInfo);
                }
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }

        private List<Sample> ReadNIMBUSWorkBook(string filePath)
        {
            // 读取NIMBUS工作清单中的板号，孔号，条码和警告信息
            var samples = new List<Sample>();
            try
            {
                IWorkbook workbook = ReadExecl(filePath);
                if (workbook == null)
                {
                    throw new IOException($"无法打开{filePath}");
                }

                ISheet worksheet = workbook.GetSheetAt(0);

                int idCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "TPositionId").ColumnIndex;
                int bcCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "SPositionBC").ColumnIndex;
                int warmCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "Warm").ColumnIndex;

                // 获取板号，最多四块板
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    var warm = row.GetCell(warmCol).ToString();
                    if (warm.StartsWith('X'))
                    {
                        Plates.Add(warm);
                    }
                }

                if (Plates.Count == 0)
                {
                    throw new ArgumentException("未识别到板号");
                }

                int p = -1;
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    if (row.Cells.Count == 0) { continue; }

                    string bc = row.GetCell(bcCol).ToString();
                    string pos = row.GetCell(idCol).ToString();
                    string warm = row.GetCell(warmCol).ToString();

                    if (pos == "A1")
                    {
                        p++;
                    }
                    string plate = Plates[p];

                    var sample = new Sample(null, bc, plate, pos, warm);
                    samples.Add(sample);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取{System.IO.Path.GetFileName(filePath)}文件时发生错误：{ex.Message}");
                throw;
            }

            return samples;
        }

        private List<Sample> ReadOperateWorkBook(List<Sample> samples, string filePath)
        {
            // 读取每日操作清单，补充实验号到samples中
            try
            {
                IWorkbook workbook = ReadExecl(filePath);
                if (workbook == null)
                {
                    throw new IOException($"无法打开{filePath}");
                }

                ISheet worksheet = workbook.GetSheetAt(0);

                int bcCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "主条码").ColumnIndex;
                int numCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "实验号").ColumnIndex;

                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    if (row.Cells.Count == 0) { continue; }
                    string barCode = row.GetCell(bcCol).ToString();
                    string number = row.GetCell(numCol).ToString();

                    var sample = samples.Find(i => i.BarCode == barCode);
                    if (sample != null)
                    {
                        sample.Number = number;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取{System.IO.Path.GetFileName(filePath)}文件时发生错误：{ex.Message}");
                throw;
            }
            return samples;
        }

        private IWorkbook ReadExecl(string filePath)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    try
                    {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                    catch
                    {
                        workbook = null;
                    }
                }
                if (workbook == null)
                {
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        workbook = new XSSFWorkbook(fileStream);
                    }
                }
            }
            catch
            {
            }
            return workbook;
        }

        class Sample
        {
            public Sample(string number, string barCode, string plate, string position, string warmInfo)
            {
                Number = number;
                if (barCode == "----------")
                {
                    barCode = "";
                }
                BarCode = barCode;
                Plate = plate;
                Position = position;
                WarmInfo = warmInfo;
            }
            public string Number { get; set; }
            public string BarCode { get; set; }
            public string Plate { get; set; }
            public string Position { get; set; }
            public string WarmInfo { get; set; }
        }

        private StackPanel Create96PlateForm(string plate)
        {
            var stack = new StackPanel();
            stack.Children.Add(Header());
            stack.Children.Add(InfoHeader(plate));
            stack.Children.Add(Body(plate));
            stack.Children.Add(Footer());
            return stack;
        }

        #region 生成96孔板工作单结构
        private Grid Body(string plate)
        {
            const int rowNum = 8;
            const int colNum = 12;

            var grid = CreateGrid();

            for (int i = 0; i <= colNum; i++)
            {
                if (i == 0)
                {
                    grid.ColumnDefinitions.Add(CreateColumn(RowTitleWidth));
                }
                else
                {
                    grid.ColumnDefinitions.Add(CreateColumn(CellColWidth));
                }
            }

            for (int i = 0; i <= rowNum; i++)
            {
                if (i == 0)
                {
                    grid.RowDefinitions.Add(CreateRow(ColumnTitleHeight));
                }
                else
                {
                    grid.RowDefinitions.Add(CreateRow(CellRowHeight));
                }
            }

            AddBorder(grid);
            AddColumnTitle(grid);
            AddRowTitle(grid);
            AddNumber(grid, plate);

            return grid;
        }

        private Grid Header()
        {
            var header = CreateGrid();

            header.RowDefinitions.Add(CreateRow(RowHeight));
            header.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth * 12));

            var txt = CreateTextBlock("____前处理样品工作单（Testing Worksheet）", 0, 0, 1, 5);
            txt.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            header.Children.Add(txt);

            AddBorder(header);

            return header;
        }

        private Grid InfoHeader(string plate)
        {
            // TODO: plateNum should read from excelfile
            var infoHeader = CreateGrid();

            infoHeader.RowDefinitions.Add(CreateRow(RowHeight));
            infoHeader.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth * 2));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth * 2));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth * 6));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth));

            infoHeader.Children.Add(CreateTextBlock("检测日期(Date)", 0, 0));
            infoHeader.Children.Add(CreateTextBlock(DateTime.Now.ToString("yyyy年MM月dd日"), 0, 1));
            infoHeader.Children.Add(CreateTextBlock($"板号：{plate}   检测仪器编号：FXS-YZ___   系统号：___", 0, 2));
            infoHeader.Children.Add(CreateTextBlock("板位", 0, 3));

            AddBorder(infoHeader);

            return infoHeader;
        }

        private Grid Footer()
        {
            Grid footer = CreateGrid();

            footer.RowDefinitions.Add(CreateRow(RowHeight));
            footer.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth * 2));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 2));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 3));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 5));

            footer.Children.Add(CreateTextBlock("加样枪：FXS-YY", 0, 0));
            footer.Children.Add(CreateTextBlock("手工加样人员：", 0, 1));
            footer.Children.Add(CreateTextBlock("加样情况审核人员：", 0, 2));
            footer.Children.Add(CreateTextBlock("备注：手写的实验号必须一一核对实验号，条码和孔位。", 0, 3));

            AddBorder(footer);

            return footer;
        }
        #endregion

        #region 生成96孔板行列标题
        private void AddColumnTitle(Grid grid)
        {
            for (int col = 1; col < 13; col++)
            {
                var txt = CreateTextBlock(col.ToString(), 0, col);
                txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                txt.VerticalAlignment = System.Windows.VerticalAlignment.Center;
                grid.Children.Add(txt);
            }
        }

        private void AddRowTitle(Grid grid)
        {
            for (int row = 1; row < 9; row++)
            {
                var txt = CreateTextBlock(((char)(row + 64)).ToString(), row, 0);
                txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                txt.VerticalAlignment = System.Windows.VerticalAlignment.Center;
                grid.Children.Add(txt);
            }
        }
        #endregion

        /// <summary>
        /// 填充实验号和条码信息
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="plate"></param>
        private void AddNumber(Grid grid, string plate)
        {
            if (Samples.Count == 0)
            {
                return;
            }

            string[] Rows = { "", "A", "B", "C", "D", "E", "F", "G", "H" };

            for (int row = 1; row < 9; row++)
            {
                for (int col = 1; col < 13; col++)
                {
                    string position = $"{Rows[row]}{col}";
                    string order = ((row - 1) * 12 + col).ToString();

                    var sample = Samples.FirstOrDefault(i => i.Plate == plate && i.Position == position);

                    if (sample != null)
                    {
                        if (sample.WarmInfo.StartsWith('X'))
                        {
                            // 定位孔
                            var txt = new TextBlock()
                            {
                                Text = sample.WarmInfo,
                                FontSize = 50,
                                FontFamily = new System.Windows.Media.FontFamily(CustomFontFamily),
                                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                                VerticalAlignment = System.Windows.VerticalAlignment.Center
                            };
                            Grid.SetColumn(txt, col);
                            Grid.SetRow(txt, row);
                            grid.Children.Add(txt);
                        }
                        else
                        {
                            string text = $"{order}   {sample?.Number} \n{sample?.BarCode}";
                            grid.Children.Add(CreateTextBlock(text, row, col));
                        }
                    }
                }
            }
        }

        private TextBlock CreateTextBlock(string text, int row, int column, int rowSpan = 1, int colSpan = 1)
        {
            var txt = new TextBlock
            {
                Text = text,
                Margin = new Thickness(5),
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new System.Windows.Media.FontFamily(CustomFontFamily),
                FontSize = 17d
            };
            Grid.SetColumn(txt, column);
            Grid.SetRow(txt, row);
            return txt;
        }

        private RowDefinition CreateRow(double height)
        {
            var row = new RowDefinition();
            row.Height = new GridLength(height);
            return row;
        }

        private ColumnDefinition CreateColumn(double width)
        {
            var col = new ColumnDefinition();
            col.Width = new GridLength(width);
            return col;
        }

        private Grid CreateGrid()
        {
            return new Grid()
            {
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };
        }

        private void AddBorder(Grid grid)
        {
            for (int row = 0; row < grid.RowDefinitions.Count; row++)
            {
                for (int col = 0; col < grid.ColumnDefinitions.Count; col++)
                {
                    var border = new Border
                    {
                        BorderThickness = new Thickness(1),
                        BorderBrush = Brushes.Black
                    };
                    Grid.SetColumn(border, col);
                    Grid.SetRow(border, row);
                    grid.Children.Add(border);
                }
            }
        }

    }
}
