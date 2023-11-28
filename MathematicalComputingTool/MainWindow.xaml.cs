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


namespace MathematicalComputingTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            _viewModel = new VMMathematicalComputingTool();
            DataContext = this._viewModel;
        }

        /// <summary>
        /// データコンテスト
        /// </summary>
        private readonly VMMathematicalComputingTool _viewModel;

        /// <summary>
        /// テキストファイルを選択するボタンがクリックされたときの処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectTxtFileButton_Click(object sender, RoutedEventArgs e)
        {
            this._viewModel.SelectFileButton_Click();
        }

        /// <summary>
        /// エクセルファイルを選択するボタンがクリックされたときの処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            this._viewModel.SelectFileButton_Click(false);
        }
    }
}
