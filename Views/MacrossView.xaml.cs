
namespace LKZ.SSMSUtils.Views
{
    using System.Windows;

    /// <summary>
    /// Interaction logic for MacrossView.xaml
    /// </summary>
    public partial class MacrossView
    {
        public MacrossView()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
