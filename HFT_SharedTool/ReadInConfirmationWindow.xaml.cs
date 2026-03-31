using System.Windows;
using System.Windows.Input;

namespace HFT_SharedTool;

public partial class ReadInConfirmationWindow : Window {
    public ReadInConfirmationWindow() {
        InitializeComponent();
    }

    public bool UserConfirmed { get; private set; }

    private void YesButton_Click(object sender, RoutedEventArgs e) {
        UserConfirmed = true;
        DialogResult = true;
        Close();
    }

    private void NoButton_Click(object sender, RoutedEventArgs e) {
        UserConfirmed = false;
        DialogResult = false;
        Close();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) {
        UserConfirmed = false;
        DialogResult = false;
        Close();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
        if (e.ButtonState == MouseButtonState.Pressed)
            DragMove();
    }
}