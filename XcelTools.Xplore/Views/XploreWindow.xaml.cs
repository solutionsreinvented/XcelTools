using System.Windows;
using System.Windows.Input;

namespace XcelTools.XpLore.Views
{
    /// <summary>
    /// Interaction logic for XploreWindow.xaml
    /// </summary>
    public partial class XploreWindow : Window
    {
        public XploreWindow()
        {
            InitializeComponent();
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}
