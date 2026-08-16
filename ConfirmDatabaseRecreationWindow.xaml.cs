using System.Windows;

namespace GNA_DLRreport
{
    public partial class ConfirmDatabaseRecreationWindow : Window
    {
        #region Constructor

        public ConfirmDatabaseRecreationWindow()
        {
            InitializeComponent();
        }

        #endregion


        #region Proceed

        private void btnProceed_Click(
            object sender,
            RoutedEventArgs e)
        {
            DialogResult = true;
        }

        #endregion


        #region Abort

        private void btnAbort_Click(
            object sender,
            RoutedEventArgs e)
        {
            DialogResult = false;
        }

        #endregion
    }
}
