using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using UpdateFieldsDocumentActivity.Commands;

namespace UpdateFieldsDocumentActivity.Message.ViewModels
{
    internal class MessageViewModel : INotifyPropertyChanged
    {
        private readonly Action<bool> _dialogResultCallback;
        private readonly Window _dialogWindow;

        public MessageViewModel(Window dialogWindow, Action<bool> dialogResultCallback = null)
        {
            _dialogWindow = dialogWindow ?? throw new ArgumentNullException(nameof(dialogWindow));
            _dialogResultCallback = dialogResultCallback;

            InitializeCommands();
        }

        public MessageViewModel() : this(null, null)
        {
            // Для дизайнера или случаев, когда окно передается позже
            InitializeCommands();
        }

        private void InitializeCommands()
        {
            OkCommand = new RelayCommand(ExecuteOkCommand, CanExecuteOkCommand);
            CancelCommand = new RelayCommand(ExecuteCancelCommand, CanExecuteCancelCommand);
        }

        private bool _isCancelButtonVisible = true;
        public bool IsCancelButtonVisible
        {
            get => _isCancelButtonVisible;
            set { _isCancelButtonVisible = value; OnPropertyChanged(nameof(IsCancelButtonVisible)); }
        }

        private string _title;
        public string Title
        {
            get => _title;
            set { _title = value; OnPropertyChanged(nameof(Title)); }
        }

        private string _content;
        public string Content
        {
            get => _content;
            set { _content = value; OnPropertyChanged(nameof(Content)); }
        }

        private ICommand _okCommand;
        public ICommand OkCommand
        {
            get => _okCommand;
            set { _okCommand = value; OnPropertyChanged(nameof(OkCommand)); }
        }

        private ICommand _cancelCommand;
        public ICommand CancelCommand
        {
            get => _cancelCommand;
            set { _cancelCommand = value; OnPropertyChanged(nameof(CancelCommand)); }
        }

        private bool CanExecuteOkCommand(object parameter)
        {
            // Можно добавить условия, когда кнопка ОК доступна
            // Например, проверка валидности введенных данных
            return true;
        }

        private void ExecuteOkCommand(object parameter)
        {
            // Закрываем окно с результатом true
            if (_dialogWindow != null)
            {
                _dialogWindow.DialogResult = true;
                _dialogWindow.Close();
            }

            // Вызываем callback, если он был передан
            _dialogResultCallback?.Invoke(true);
        }

        private bool CanExecuteCancelCommand(object parameter)
        {
            // Кнопка отмены всегда доступна
            return true;
        }

        private void ExecuteCancelCommand(object parameter)
        {
            // Закрываем окно с результатом false
            if (_dialogWindow != null)
            {
                _dialogWindow.DialogResult = false;
                _dialogWindow.Close();
            }

            // Вызываем callback, если он был передан
            _dialogResultCallback?.Invoke(false);
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
