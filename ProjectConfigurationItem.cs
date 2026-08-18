using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GNA_DLRreport
{
    #region Project Configuration Item

    public sealed class ProjectConfigurationItem : INotifyPropertyChanged
    {
        #region Private Fields

        private string _projectName = string.Empty;
        private bool _isDeleted;
        private bool _isActive;

        #endregion


        #region Database Properties

        public int Project_ID { get; init; }

        public string ProjectName
        {
            get
            {
                return _projectName;
            }

            set
            {
                string newValue =
                    value
                    ?? throw new ArgumentNullException(
                        paramName: nameof(value));

                if (_projectName == newValue)
                {
                    return;
                }

                _projectName = newValue;

                OnPropertyChanged();
            }
        }

        public bool IsDeleted
        {
            get
            {
                return _isDeleted;
            }

            set
            {
                if (_isDeleted == value)
                {
                    return;
                }

                _isDeleted = value;

                OnPropertyChanged();
            }
        }

        #endregion


        #region Application Properties

        public bool IsActive
        {
            get
            {
                return _isActive;
            }

            set
            {
                if (_isActive == value)
                {
                    return;
                }

                _isActive = value;

                OnPropertyChanged();
            }
        }

        public string OriginalProjectName { get; set; } = string.Empty;

        #endregion


        #region Property Change Notification

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(
            [CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(
                sender: this,
                e: new PropertyChangedEventArgs(
                    propertyName: propertyName));
        }

        #endregion
    }

    #endregion
}
