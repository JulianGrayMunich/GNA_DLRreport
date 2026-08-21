#region System Preparation

using System;
using System.Windows;

#endregion

namespace GNA_DLRreport
{
    public partial class ArrayDetailsWindow : Window
    {
        #region Result State

        public bool DeleteRequested { get; private set; }

        #endregion


        #region Constructor

        public ArrayDetailsWindow(
            int arrayId,
            string arrayName,
            string arrayTypeName,
            string pointA,
            string pointB,
            string pointC,
            string pointD,
            string pointE)
        {
            #region Initialise Window Components

            InitializeComponent();

            #endregion


            #region Validate Array Details

            if (arrayId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(arrayId),
                    message: "Array_ID must be greater than zero.");
            }

            #endregion


            #region Populate Array Details

            txtArrayId.Text =
                arrayId.ToString();

            txtArrayName.Text =
                arrayName ?? string.Empty;

            txtArrayType.Text =
                arrayTypeName ?? string.Empty;

            txtPointA.Text =
                pointA ?? string.Empty;

            txtPointB.Text =
                pointB ?? string.Empty;

            txtPointC.Text =
                pointC ?? string.Empty;

            txtPointD.Text =
                pointD ?? string.Empty;

            txtPointE.Text =
                pointE ?? string.Empty;

            #endregion
        }

        #endregion


        #region Dialog Actions

        private void btnRetainArray_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Retain Array

            DeleteRequested =
                false;

            DialogResult =
                true;

            #endregion
        }


        private void btnDeleteArray_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Request Array Delete

            DeleteRequested =
                true;

            DialogResult =
                true;

            #endregion
        }

        #endregion
    }
}
