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
using static TestDocx.ExcelParser;

namespace BuBulls
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MessageBox : Window
    {
        List<ParsingError> _viewModel;
        public MessageBox()
        {
            InitializeComponent();            
            
        }

        public void Show(List<ParsingError> _viewModel)
        {
            
            this._viewModel = _viewModel;
            this.DataContext = _viewModel;
            this.lvErrors.ItemsSource = _viewModel;
            this.Show();
        }

        private void btn_onClose(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
