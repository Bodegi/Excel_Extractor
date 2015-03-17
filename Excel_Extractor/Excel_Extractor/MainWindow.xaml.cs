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
using System.IO;
using Extraction;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace Excel_Extractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        
        private void btnBrowseIn_Click(object sender, RoutedEventArgs e)
        {
            var open = new System.Windows.Forms.FolderBrowserDialog();
            var fileInteraction = open.ShowDialog();
            switch (fileInteraction)
            {
                case System.Windows.Forms.DialogResult.OK:
                    string path = open.SelectedPath;
                    txtDir.Text = path;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtDir.Text = "Select Directory to Search";
                    break;
            }

        }

        private void btnBrowseOut_Click(object sender, RoutedEventArgs e)
        {
            var open = new System.Windows.Forms.FolderBrowserDialog();
            var fileInteraction = open.ShowDialog();
            switch (fileInteraction)
            {
                case System.Windows.Forms.DialogResult.OK:
                    string path = open.SelectedPath;
                    txtOut.Text = path;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtOut.Text = "Select Location for Output";
                    break;
            }

        }

        private void btnExtract_Click(object sender, RoutedEventArgs e)
        {
            if(txtDir.Text != "Select Directory to Search")
            {
                if(txtOut.Text != "Select Location for Output")
                {
                    if (txtFileName.Text != null)
                    {
                        List<string> visited = new List<string>();
                        string output = txtOut.Text + "\\" + txtFileName.Text + ".XLSX";
                        FileSearch.traversal(txtDir.Text, visited, output, true);
                    }
                    else
                    {
                        MessageBox.Show("Please select a name for the completed workbook", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("Please select a location to output the final workbook", "Error");
                }
            }
            else
            {
                MessageBox.Show("Please select a directory to search", "Error");
            }
        }
    }
}
