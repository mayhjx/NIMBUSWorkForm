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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const double ColumnTitleHeight = 25d; // 96孔板列标题行高
        private const double RowTitleWidth = 25d; // 96孔板行标题列宽
        private const double CellRowHeight = 50d; // 96孔板单元格行高
        private const double CellColWidth = 90d;  // 96孔板单元格列宽
        private const double RowHeight = 30d; // 工作清单header和footer行高

        public MainWindow()
        {
            InitializeComponent();
            PrintArea.Children.Add(Create96PlateForm());
            PrintArea.Children.Add(Create96PlateForm());
            PrintArea.Children.Add(Create96PlateForm());
            PrintArea.Children.Add(Create96PlateForm());
        }

        private StackPanel Create96PlateForm()
        {
            var stack = new StackPanel();
            stack.Children.Add(Header());
            stack.Children.Add(InfoHeader());
            stack.Children.Add(Body());
            stack.Children.Add(Footer());
            return stack;
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog dialog = new PrintDialog();

            if (dialog.ShowDialog().GetValueOrDefault() != true)
            {
                return;
            }

            dialog.PrintVisual(PrintArea.Children[0], "print test");
        }

        private Grid Body()
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
            AddNumber(grid);

            return grid;
        }

        private Grid Header()
        {
            var header = CreateGrid();

            header.RowDefinitions.Add(CreateRow(RowHeight));
            header.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth * 12));

            header.Children.Add(CreateTextBlock("前处理样品工作单（Testing Worksheet）", 0, 0, 1, 5));

            AddBorder(header);

            return header;
        }

        private Grid InfoHeader(int plateNum = 1)
        {
            // TODO: plateNum should read from excelfile
            var infoHeader = CreateGrid();

            infoHeader.RowDefinitions.Add(CreateRow(RowHeight));
            infoHeader.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth * 2));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth * 7));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth));
            infoHeader.ColumnDefinitions.Add(CreateColumn(CellColWidth));

            infoHeader.Children.Add(CreateTextBlock("检测日期（Date)", 0, 0));
            infoHeader.Children.Add(CreateTextBlock(DateTime.Now.ToShortDateString(), 0, 1));
            infoHeader.Children.Add(CreateTextBlock($"板号：{plateNum} 检测仪器编号：FXS-YZ__ 系统号：__", 0, 2));
            infoHeader.Children.Add(CreateTextBlock("板位：", 0, 3));

            AddBorder(infoHeader);

            return infoHeader;
        }

        private Grid Footer()
        {
            Grid footer = CreateGrid();
            footer.Height = 80; // 为了和下一个Grid空开

            footer.RowDefinitions.Add(CreateRow(RowHeight));
            footer.ColumnDefinitions.Add(CreateColumn(RowTitleWidth + CellColWidth * 2));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 2));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 3));
            footer.ColumnDefinitions.Add(CreateColumn(CellColWidth * 5));

            footer.Children.Add(CreateTextBlock("加样枪：FXS-YY", 0, 0));
            footer.Children.Add(CreateTextBlock("手工加样人员：", 0, 1));
            footer.Children.Add(CreateTextBlock("加样情况审核人员：", 0, 2));
            footer.Children.Add(CreateTextBlock("备注：手写的实验号必须一一核对实验号，条码和孔位", 0, 3));

            AddBorder(footer);

            return footer;
        }

        private Grid CreateGrid()
        {
            return new Grid()
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
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

        private TextBlock CreateTextBlock(string text, int row, int column, int rowSpan = 1, int colSpan = 1)
        {
            var txt = new TextBlock
            {
                Text = text,
                Padding = new Thickness(5)
            };
            Grid.SetColumn(txt, column);
            Grid.SetRow(txt, row);
            return txt;
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

        #region 生成96孔板行列标题
        private void AddColumnTitle(Grid grid)
        {
            for (int col = 1; col < 13; col++)
            {
                var t = new TextBlock
                {
                    Text = col.ToString(),
                    FontSize = 20,
                    TextAlignment = TextAlignment.Center
                };
                Grid.SetColumn(t, col);
                Grid.SetRow(t, 0);
                grid.Children.Add(t);
            }
        }

        private void AddRowTitle(Grid grid)
        {
            for (int row = 1; row < 9; row++)
            {
                var t = new TextBlock
                {
                    Text = ((char)(row + 64)).ToString(),
                    FontSize = 20,
                    TextAlignment = TextAlignment.Center
                };
                Grid.SetRow(t, row);
                Grid.SetColumn(t, 0);
                grid.Children.Add(t);
            }
        }
        #endregion

        private void AddNumber(Grid grid)
        {
            // TODO 根据移液清单中的条码和孔号生成
            for (int row = 1; row < 9; row++)
            {
                for (int col = 1; col < 13; col++)
                {
                    string num = ((row - 1) * 12 + col).ToString();
                    grid.Children.Add(CreateTextBlock(num, row, col));
                }
            }
        }
    }
}
