using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace NIMBUSWorkForm
{
    /// <summary>
    /// Interaction logic for Setting.xaml
    /// </summary>
    public partial class Setting : Window
    {
        public Setting()
        {
            InitializeComponent();
            ListItem();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.Save();
            this.Close();
        }

        private void ListItem()
        {
            TargetPreFixList.ItemsSource = null;
            TargetPreFixList.ItemsSource = Settings.Default.TargetPreFix;
            TargetPreFixList.SelectedIndex = -1;
        }

        private void TargetPreFixList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(TargetPreFixList.SelectedItem != null)
            {
                SelectedItem.Text = TargetPreFixList.SelectedItem.ToString();
            }
            else
            {
                SelectedItem.Clear();
            }
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            var text = SelectedItem.Text.Trim();
            if(!string.IsNullOrEmpty(text))
            {
                Settings.Default.TargetPreFix.Add(text);
                Settings.Default.Save();
                SelectedItem.Clear();
                SelectedItem.Focus();
                ListItem();
            }
        }

        private void DeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            var text = SelectedItem.Text.Trim();
            if (!string.IsNullOrEmpty(text))
            {
                Settings.Default.TargetPreFix.Remove(text);
                Settings.Default.Save();
                SelectedItem.Clear();
                SelectedItem.Focus();
                ListItem();
            }
        }
    }
}
