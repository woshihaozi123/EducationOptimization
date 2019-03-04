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

using System.Windows.Shapes;

using System.Windows.Forms;

using System.Data;
using EducOpti.source;
namespace EducOpti
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_District_Click(object sender, RoutedEventArgs e)
        {
            
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "输入文件";
            openFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "xls";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string inputFile = openFileDialog.FileName;
            this.DistrictTextBox.Text = inputFile;

        
        }

        private void btn_School_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "输入文件";
            openFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "xls";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel) { return; }

            string inputFile = openFileDialog.FileName;
            this.SchoolTextBox.Text = inputFile;
        }

        private void btn_Relation_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "输入文件";
            openFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "xls";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel) { return; }

            string inputFile = openFileDialog.FileName;
            this.RelationTextBox.Text = inputFile;
          
        }

        private void btn_Output_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "输出文件";
            saveFileDialog.Filter = "excel文件|*.xls|excel文件|*.xlsx|所有文件|*.*";
            saveFileDialog.FileName = string.Empty;
            DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string outputFile = saveFileDialog.FileName;
            this.OutputTextBox.Text = outputFile;
        }

        private void btOk_Click(object sender, RoutedEventArgs e)
        {
            if (this.DistrictTextBox.Text != string.Empty && this.SchoolTextBox.Text != string.Empty && this.RelationTextBox.Text != string.Empty && this.OutputTextBox.Text != string.Empty)
            {
                string districtFilePath = this.DistrictTextBox.Text;
                string schoolFilePath = this.SchoolTextBox.Text;
                string relationFilePath = this.RelationTextBox.Text;
                string outputFilePath = this.OutputTextBox.Text;
                EducOptiShp eo = new EducOptiShp();
                DataTable a = eo.ExcelToDataTable(districtFilePath, true);
                DataTable b = eo.ExcelToDataTable(schoolFilePath, true);

                DataTable c = eo.ExcelToDataTable(relationFilePath, true);
               DataTable d = eo.CalNewRelation(a, b, c);
               eo.DataTableToExcel(d, outputFilePath);
               System.Windows.Forms.MessageBox.Show("输出完毕！");
            }
            else
                System.Windows.Forms.MessageBox.Show("路径不能为空！！！");
        }

        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        
    }
}
