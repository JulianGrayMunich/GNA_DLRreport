#region System Preparation

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

#endregion

namespace GNA_DLRreport
{
    public sealed class ArraySelectionItem
    {
        public int Array_ID { get; init; }

        public string ArrayTypeName { get; init; } =
            string.Empty;

        public string ArrayName { get; init; } =
            string.Empty;
    }


    public partial class ArraySelectionWindow : Window
    {
        #region Selected Array

        public int? SelectedArrayId { get; private set; }

        #endregion


        #region Constructor

        public ArraySelectionWindow(
            string projectName,
            IReadOnlyList<ArraySelectionItem> arrays)
        {
            #region Initialise Window

            InitializeComponent();

            #endregion


            #region Validate Parameters

            string validatedProjectName =
                projectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName:
                        nameof(projectName));

            if (string.IsNullOrWhiteSpace(
                value:
                    validatedProjectName))
            {
                throw new ArgumentException(
                    message:
                        "Project name required.",
                    paramName:
                        nameof(projectName));
            }

            ArgumentNullException.ThrowIfNull(
                argument:
                    arrays);

            #endregion


            #region Populate Committed Arrays

            txtProjectName.Text =
                validatedProjectName;

            dgCommittedArrays.ItemsSource =
                arrays;

            btnViewSelectedArray.IsEnabled =
                false;

            #endregion
        }

        #endregion


        #region Array Selection

        private void dgCommittedArrays_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Update View Button

            btnViewSelectedArray.IsEnabled =
                dgCommittedArrays.SelectedItem
                    is ArraySelectionItem;

            #endregion
        }


        private void dgCommittedArrays_MouseDoubleClick(
            object sender,
            MouseButtonEventArgs e)
        {
            #region View Double-Clicked Array

            CommitSelectedArray();

            #endregion
        }


        private void btnViewSelectedArray_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region View Selected Array

            CommitSelectedArray();

            #endregion
        }


        private void CommitSelectedArray()
        {
            #region Validate Selected Array

            if (dgCommittedArrays.SelectedItem
                is not ArraySelectionItem selectedArray)
            {
                return;
            }

            #endregion


            #region Return Selected Array

            SelectedArrayId =
                selectedArray.Array_ID;

            DialogResult =
                true;

            Close();

            #endregion
        }

        #endregion


        #region Cancel Selection

        private void btnCancel_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Cancel Array Selection

            DialogResult =
                false;

            Close();

            #endregion
        }

        #endregion
    }
}
