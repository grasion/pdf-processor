using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using QRCoder;

namespace PDFProcessor
{
    public partial class MainWindow
    {
        private const string CoffeeSupportUrl = "https://qr.kakaopay.com/Ej9Q6cblA9c409206";

        private void CoffeeBanner_Click(object sender, MouseButtonEventArgs e)
        {
            try { CreateCoffeePopup().ShowDialog(); }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"후원 QR 코드를 표시할 수 없습니다.\n{ex.Message}",
                    "커피 사주기", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private Window CreateCoffeePopup()
        {
            var popup = new Window
            {
                Title = "개발자에게 커피 사주기", Width = 380, SizeToContent = SizeToContent.Height,
                Owner = this, WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize, Background = Brushes.White, FontFamily = FontFamily
            };
            var panel = new StackPanel { Margin = new Thickness(24) };
            panel.Children.Add(new TextBlock
            {
                Text = "☕ 개발자에게 커피 사주기", FontSize = 18, FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(168, 91, 19)),
                HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 18)
            });
            using var generator = new QRCodeGenerator();
            using var data = generator.CreateQrCode(CoffeeSupportUrl, QRCodeGenerator.ECCLevel.Q);
            using var code = new QRCode(data);
            using var bitmap = code.GetGraphic(12);
            panel.Children.Add(new Image
            {
                Source = ConvertBitmapToBitmapSource(bitmap), Width = 280, Height = 280,
                Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Center
            });
            panel.Children.Add(new TextBlock
            {
                Text = "QR 코드를 스캔하여 개발자에게\n커피 한 잔의 응원을 보내주세요.",
                TextAlignment = TextAlignment.Center, FontSize = 13, Foreground = Brushes.DimGray,
                Margin = new Thickness(0, 16, 0, 18)
            });
            var close = new Button { Content = "닫기", MinWidth = 90, Padding = new Thickness(16, 6, 16, 6),
                HorizontalAlignment = HorizontalAlignment.Center, IsCancel = true };
            close.Click += (sender, args) => popup.Close();
            panel.Children.Add(close);
            popup.Content = panel;
            return popup;
        }
    }
}
