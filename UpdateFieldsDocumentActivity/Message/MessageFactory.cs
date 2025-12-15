using Ascon.Pilot.Theme.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using UpdateFieldsDocumentActivity.Message.ViewModels;
using UpdateFieldsDocumentActivity.Message.Views;

namespace UpdateFieldsDocumentActivity
{
    internal class MessageFactory : DialogWindowFactory
    {
        // Метод для простого сообщения с кнопками ОК/Отмена
        public static DialogWindow CreateDialog(string content, bool cancel = true, string title = null)
        {
            var dialogWindow = new DialogWindow();

            var viewModel = new MessageViewModel(dialogWindow);
            viewModel.Title = title ?? "Сообщение";
            viewModel.Content = content;
            viewModel.IsCancelButtonVisible = cancel;

            dialogWindow.SourceInitialized += FixLayoutObjectSourceInitialized;
            dialogWindow.LayoutUpdated += FixLayoutObjectLayoutUpdated;
            dialogWindow.Title = viewModel.Title;
            dialogWindow.SizeToContent = SizeToContent.WidthAndHeight;

            var messageView = new MessageView();
            dialogWindow.DataContext = viewModel;
            dialogWindow.Content = messageView;
            dialogWindow.Owner = Application.Current.MainWindow;
            dialogWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            return dialogWindow;
        }

        // Метод для сообщения с callback обработкой результата
        public static DialogWindow CreateDialogWithCallback(
            string content,
            Action<bool> resultCallback,
            string title = null,
            bool showCancel = true)
        {
            var dialogWindow = new DialogWindow();

            var viewModel = new MessageViewModel(dialogWindow, resultCallback);
            viewModel.Title = title ?? "Сообщение";
            viewModel.Content = content;
            viewModel.IsCancelButtonVisible = showCancel;

            dialogWindow.SourceInitialized += FixLayoutObjectSourceInitialized;
            dialogWindow.LayoutUpdated += FixLayoutObjectLayoutUpdated;
            dialogWindow.Title = viewModel.Title;
            dialogWindow.SizeToContent = SizeToContent.WidthAndHeight;

            var messageView = new MessageView();
            dialogWindow.DataContext = viewModel;
            dialogWindow.Content = messageView;
            dialogWindow.Owner = Application.Current.MainWindow;
            dialogWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            return dialogWindow;
        }

        // Метод для сообщения только с кнопкой ОК
        public static DialogWindow CreateInfoDialog(string content, string title = null)
        {
            var dialogWindow = new DialogWindow();

            var viewModel = new MessageViewModel(dialogWindow);
            viewModel.Title = title ?? "Информация";
            viewModel.Content = content;
            viewModel.IsCancelButtonVisible = false;

            dialogWindow.SourceInitialized += FixLayoutObjectSourceInitialized;
            dialogWindow.LayoutUpdated += FixLayoutObjectLayoutUpdated;
            dialogWindow.Title = viewModel.Title;
            dialogWindow.SizeToContent = SizeToContent.WidthAndHeight;

            var messageView = new MessageView();
            dialogWindow.DataContext = viewModel;
            dialogWindow.Content = messageView;
            dialogWindow.Owner = Application.Current.MainWindow;
            dialogWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            return dialogWindow;
        }

        // Метод для подтверждения действия
        public static DialogWindow CreateConfirmationDialog(
            string question,
            Action<bool> confirmationCallback,
            string title = "Подтверждение")
        {
            return CreateDialogWithCallback(question, confirmationCallback, title, true);
        }

        // Статические методы для удобного использования
        public static bool? ShowDialog(string content, bool cancel = true, string title = null)
        {
            var dialog = CreateDialog(content, cancel, title);
            return dialog.ShowDialog();
        }

        public static void ShowInfo(string content, string title = null)
        {
            var dialog = CreateInfoDialog(content, title);
            dialog.ShowDialog();
        }

        public static bool ShowConfirmation(string question, string title = "Подтверждение")
        {
            var dialog = CreateDialog(question, true, title);
            return dialog.ShowDialog() == true;
        }

        public static void ShowDialogWithCallback(
            string content,
            Action<bool> resultCallback,
            string title = null,
            bool showCancel = true)
        {
            var dialog = CreateDialogWithCallback(content, resultCallback, title, showCancel);
            dialog.ShowDialog();
        }
    }
}
