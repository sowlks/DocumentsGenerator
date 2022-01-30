using System.Windows;

namespace Demonstration
{
    public partial class MainWindow : Window
    {
        private readonly MainWindowUI context = new ();

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = context;
        }
    }
}
