using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;
using Task = System.Threading.Tasks.Task;

namespace WC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Application _application;

        public MainWindow()
        {
            InitializeComponent();
            _application = new Application();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            ((Button) sender).Content = "Пачекайтэ...";
            ((Button) sender).IsEnabled = false;

            await DoJob(((Button)sender));
        }

        private async Task DoJob(ContentControl sender)
        {
            var dialog = new OpenFileDialog { Filter = "Word (*.doc,*.docx)|*.doc;*.docx" };

            if (dialog.ShowDialog() ?? false)
            {
                var doc = _application.Documents.Open(dialog.FileName, null, true);

                var text = doc.Words.OfType<Range>()
                    .Select(x => x.Text?.ToLower().Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x) && char.IsLetterOrDigit(x.First()))
                    .Distinct()
                    .ToList();

                await File.WriteAllTextAsync($"{dialog.FileName}.txt", string.Join(Environment.NewLine, text));

                MessageBox.Show($"Уникальных слов: {text.Count}. Список слов сохранен: {dialog.FileName}.txt");
                
                doc.Close();
                sender.Content = "Жми меня";
                sender.IsEnabled = true;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            _application.Quit();
            base.OnClosing(e);
        }
    }
}
