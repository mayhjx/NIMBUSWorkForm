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
        public MainWindow()
        {
            InitializeComponent();
            Create96PlateForm("P1");
            Create96PlateForm("P2");
            Create96PlateForm("P3");
            Create96PlateForm("P4");
        }


        private void Create96PlateForm(string name)
        {
            var stack = new StackPanel();

            stack.Children.Add(AddHeader());
            stack.Children.Add(AddInfoHeader());
            stack.Children.Add(CreateGrid());
            stack.Children.Add(AddFooter());
            PrintArea.Children.Add(stack);

        }
        private Grid CreateGrid()
        {
            var grid = new Grid();
            grid.HorizontalAlignment = HorizontalAlignment.Center;
            grid.VerticalAlignment = VerticalAlignment.Center;
            grid.Width = 1105;
            grid.Height = 425;

            for (int i = 0; i < 13; i++)
            {
                var col = new ColumnDefinition();
                if (i == 0)
                {
                    col.Width = new GridLength(25);
                }
                else
                {
                    col.Width = new GridLength(90);
                }
                grid.ColumnDefinitions.Add(col);
            }

            for (int i = 0; i < 9; i++)
            {
                var row = new RowDefinition();
                if (i == 0)
                {
                    row.Height = new GridLength(25);
                }
                else
                {
                    row.Height = new GridLength(50);
                }
                grid.RowDefinitions.Add(row);
            }

            AddColumnTitle(grid);
            AddRowTitle(grid);
            AddBorder(grid);
            AddNumber(grid);
            return grid;
        }

        private Grid AddHeader()
        {
            var header = new Grid();
            header.HorizontalAlignment = HorizontalAlignment.Center;
            header.VerticalAlignment = VerticalAlignment.Center;

            var row = new RowDefinition();
            row.Height = new GridLength(30);
            header.RowDefinitions.Add(row);

            var column = new ColumnDefinition();
            column.Width = new GridLength(1105);
            header.ColumnDefinitions.Add(column);

            var mainTitle = new TextBlock();
            mainTitle.Text = "前处理样品工作单（Testing Worksheet）";
            Grid.SetRow(mainTitle, 0);
            Grid.SetColumn(mainTitle, 0);
            Grid.SetColumnSpan(mainTitle, 5);

            header.Children.Add(mainTitle);

            AddBorder(header);
            return header;
        }


        private Grid AddInfoHeader(int plateNum = 1)
        {
            // TODO: plateNum should read from excelfile

            var header = new Grid();
            header.HorizontalAlignment = HorizontalAlignment.Center;
            header.VerticalAlignment = VerticalAlignment.Center;

            var row = new RowDefinition();
            row.Height = new GridLength(30);
            header.RowDefinitions.Add(row);

            var column1 = new ColumnDefinition();
            column1.Width = new GridLength(115);
            header.ColumnDefinitions.Add(column1);

            var column2 = new ColumnDefinition();
            column2.Width = new GridLength(180);
            header.ColumnDefinitions.Add(column2);

            var column3 = new ColumnDefinition();
            column3.Width = new GridLength(630);
            header.ColumnDefinitions.Add(column3);

            var column4 = new ColumnDefinition();
            column4.Width = new GridLength(90);
            header.ColumnDefinitions.Add(column4);

            var column5 = new ColumnDefinition();
            column5.Width = new GridLength(90);
            header.ColumnDefinitions.Add(column5);


            var dateTxt = new TextBlock();
            dateTxt.Text = "检测日期（Date)";
            Grid.SetRow(dateTxt, 0);
            Grid.SetColumn(dateTxt, 0);
            header.Children.Add(dateTxt);

            var date = new TextBlock();
            date.Text = DateTime.Now.ToShortDateString();
            Grid.SetRow(date, 0);
            Grid.SetColumn(date, 1);
            header.Children.Add(date);

            var place = new TextBlock();
            place.Text = $"板号：{plateNum} 检测仪器编号：FXS-YZ__ 系统号：__";
            Grid.SetRow(place, 0);
            Grid.SetColumn(place, 2);
            header.Children.Add(place);

            var position = new TextBlock();
            position.Text = "板位：";
            Grid.SetRow(position, 0);
            Grid.SetColumn(position, 3);
            header.Children.Add(position);

            AddBorder(header);
            return header;
        }


        private Grid AddFooter()
        {
            var footer = new Grid();
            footer.Height = 100;
            footer.HorizontalAlignment = HorizontalAlignment.Center;
            footer.VerticalAlignment = VerticalAlignment.Center;

            var row = new RowDefinition();
            row.Height = new GridLength(30);
            footer.RowDefinitions.Add(row);

            var column1 = new ColumnDefinition();
            column1.Width = new GridLength(205);
            footer.ColumnDefinitions.Add(column1);

            var column2 = new ColumnDefinition();
            column2.Width = new GridLength(180);
            footer.ColumnDefinitions.Add(column2);

            var column3 = new ColumnDefinition();
            column3.Width = new GridLength(270);
            footer.ColumnDefinitions.Add(column3);

            var column4 = new ColumnDefinition();
            column4.Width = new GridLength(450);
            footer.ColumnDefinitions.Add(column4);

            var pipeTxt = new TextBox();
            pipeTxt.Text = "加样枪：FXS-YY";
            Grid.SetRow(pipeTxt, 0);
            Grid.SetColumn(pipeTxt, 0);
            footer.Children.Add(pipeTxt);

            var operatorTxt = new TextBox();
            operatorTxt.Text = "手工加样人员：";
            Grid.SetRow(operatorTxt, 0);
            Grid.SetColumn(operatorTxt, 1);
            footer.Children.Add(operatorTxt);

            var reviewTxt = new TextBox();
            reviewTxt.Text = "加样情况审核人员：";
            Grid.SetRow(reviewTxt, 0);
            Grid.SetColumn(reviewTxt, 2);
            footer.Children.Add(reviewTxt);

            var remark = new TextBox();
            remark.Text = "备注：手写的实验号必须一一核对实验号，条码和孔位";
            Grid.SetRow(remark, 0);
            Grid.SetColumn(remark, 3);
            footer.Children.Add(remark);

            AddBorder(footer);
            return footer;
        }


        private void AddColumnTitle(Grid grid)
        {
            for (int col = 1; col < 13; col++)
            {
                var t = new TextBlock();
                t.Text = col.ToString();
                t.FontSize = 20;
                t.TextAlignment = TextAlignment.Center;
                Grid.SetColumn(t, col);
                Grid.SetRow(t, 0);
                grid.Children.Add(t);
            }
        }

        private void AddRowTitle(Grid grid)
        {
            for (int row = 1; row < 9; row++)
            {
                var t = new TextBlock();
                t.Text = ((char)(row + 64)).ToString();
                t.FontSize = 20;
                t.TextAlignment = TextAlignment.Center;
                Grid.SetRow(t, row);
                Grid.SetColumn(t, 0);
                grid.Children.Add(t);
            }
        }

        private void AddBorder(Grid grid)
        {
            for (int row = 0; row < grid.RowDefinitions.Count; row++)
            {
                for (int col = 0; col < grid.ColumnDefinitions.Count; col++)
                {
                    var border = new Border();
                    border.BorderThickness = new Thickness(1);
                    border.BorderBrush = Brushes.Black;
                    Grid.SetColumn(border, col);
                    Grid.SetRow(border, row);
                    grid.Children.Add(border);
                }
            }
        }

        private void AddNumber(Grid grid)
        {
            for (int row = 1; row < 9; row++)
            {
                for (int col = 1; col < 13; col++)
                {
                    var num = new TextBlock();
                    num.Text = ((row - 1) * 12 + col).ToString();
                    num.Text += " Sample";
                    num.Text += " \n123456789";
                    num.Padding = new Thickness(5);
                    Grid.SetColumn(num, col);
                    Grid.SetRow(num, row);
                    grid.Children.Add(num);
                }
            }
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
    }
}
