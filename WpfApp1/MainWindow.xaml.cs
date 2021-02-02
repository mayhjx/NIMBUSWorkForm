using Microsoft.Win32;
using NIMBUSWorkForm;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string CustomFontFamily = "Times New Roman, 宋体";
        private const double ColumnTitleHeight = 6d * 3.7795275591; // 96孔板列标题行高(mm2px)
        private const double RowTitleWidth = 6d * 3.7795275591; // 96孔板行标题列宽(mm2px)
        private const double CellRowHeight = 22d * 3.7795275591; // 96孔板单元格行高(mm2px)
        private const double CellColWidth = 24d * 3.7795275591;  // 96孔板单元格列宽(mm2px)
        private const double RowHeight = 7d * 3.7795275591; // 工作清单header和footer行高(mm2px)

        private SampleTable SampleTable;

        public MainWindow()
        {
            InitializeComponent();
            this.MinWidth = RowTitleWidth + CellColWidth * 12 + 50;
            this.MinHeight = ColumnTitleHeight + CellRowHeight * 8 + RowHeight * 3 + 70;
            DocumentPage.Width = RowTitleWidth + CellColWidth * 12 + 50;
            DocumentPage.Height = ColumnTitleHeight + CellRowHeight * 8 + RowHeight * 3 + 70;
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
            SampleTable = new SampleTable();
            DocumentPage.Document = null;

            try
            {
                // 读取数据
                SampleTable = ReadUploadedExcel.ReadNIMBUSWorkBook(openFileDialog.FileName);
                SampleTable = ReadUploadedExcel.ReadOperateWorkBook(SampleTable, openFileDialog1.FileName);

                // 生成96孔板工作清单
                var flowDocument = new FlowDocument()
                {
                    ColumnWidth = RowTitleWidth + CellColWidth * 12 + 50,
                };
                foreach (var plate in SampleTable.PlateNumber)
                {
                    flowDocument.Blocks.Add(new BlockUIContainer(Create96WellPlateForm(plate)));
                }
                DocumentPage.Document = flowDocument;

                PrintButton.IsEnabled = true;
                OutputButton.IsEnabled = true;
            }
            catch (IOException ex)
            {
                MessageBox.Show($"读取Excel文件时发成错误：{ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成96孔板工作清单时出现错误：{ex.Message}");
                return;
            }
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
                BatchAndWarnFile.Create(SampleTable, filePath);
            }
        }

        /// <summary>
        /// 填充实验号和条码信息
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="plate"></param>
        private void AddNumber(Grid grid, string plate)
        {
            if (SampleTable.Samples.Count == 0)
            {
                return;
            }

            string[] Rows = { "", "A", "B", "C", "D", "E", "F", "G", "H" };

            for (int row = 1; row < 9; row++)
            {
                for (int col = 1; col < 13; col++)
                {
                    string position = $"{Rows[row]}{col}";
                    var sample = SampleTable.Samples.FirstOrDefault(i => i.Plate == plate && i.Position == position);

                    if (sample != null)
                    {
                        if (sample.IsX())
                        {
                            // 定位孔
                            var txt = new TextBlock()
                            {
                                Text = sample.Number,
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
                            var txt = new TextBlock()
                            {
                                Margin = new Thickness(5),
                                TextWrapping = TextWrapping.Wrap,
                                FontSize = 16,
                                FontFamily = new System.Windows.Media.FontFamily(CustomFontFamily),
                            };
                            txt.Inlines.Add(new Run($"{sample?.Order}\n") { FontSize = 12 });
                            txt.Inlines.Add(new Run($"{sample?.Number}\n") { FontWeight = FontWeights.Bold });
                            txt.Inlines.Add(new Run($"{sample?.BarCode}"));

                            Grid.SetColumn(txt, col);
                            Grid.SetRow(txt, row);
                            grid.Children.Add(txt);
                        }
                    }
                }
            }
        }

        #region 生成96孔板工作单结构
        private StackPanel Create96WellPlateForm(string plate)
        {
            var form = new StackPanel();
            form.Children.Add(Header());
            form.Children.Add(InfoHeader(plate));
            form.Children.Add(Body(plate));
            form.Children.Add(Footer());
            return form;
        }

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

        #region Helper Methods
        private TextBlock CreateTextBlock(string text, int row, int column, int rowSpan = 1, int colSpan = 1)
        {
            var txt = new TextBlock
            {
                Text = text,
                Margin = new Thickness(5),
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new System.Windows.Media.FontFamily(CustomFontFamily),
                FontSize = 14d
            };
            Grid.SetColumn(txt, column);
            Grid.SetRow(txt, row);
            return txt;
        }

        private RowDefinition CreateRow(double height)
        {
            var row = new RowDefinition();
            row.Height = new GridLength(height, GridUnitType.Pixel);
            return row;
        }

        private ColumnDefinition CreateColumn(double width)
        {
            var col = new ColumnDefinition();
            col.Width = new GridLength(width, GridUnitType.Pixel);
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
        #endregion
    }
}
