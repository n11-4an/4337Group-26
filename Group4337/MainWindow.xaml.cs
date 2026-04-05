using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void ShowInfo_Click(object sender, RoutedEventArgs e)
        {
            Gilmanova infoWindow = new Gilmanova();
            infoWindow.InfoText.Text = "Гильманова Азиза\nВозраст: 18\nГруппа: 4337";
            infoWindow.ShowDialog();
        }
    }
}