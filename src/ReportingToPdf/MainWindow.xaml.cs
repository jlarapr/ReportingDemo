using System.Data;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using ReportingToPdf.Models;
using ReportingToPdf.Services;
using ReportingToPdf.tools;

namespace ReportingToPdf;

/// <summary>
///     Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly IToolsService _tool;

    public MainWindow()
    {
        InitializeComponent();
        _tool = new ToolsService();
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Task.Run(async () => await RunReports()).Wait();

            MessageBox.Show("Exporting to PDF", "Done", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task RunReports()
    {
        const string cartaFilePage1 = "Assets\\CartaReport.trdp";
        const string cartaFilePage2 = "Assets\\CartaReportPage2.trdp";

        if (!File.Exists(cartaFilePage1))
        {
            MessageBox.Show("CartaReport.trdp file not found", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!File.Exists(cartaFilePage2))
        {
            MessageBox.Show("CartaReportPage2.trdp file not found", "Error", MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        ICreateReportService reportService = new CreateReportService();

        OpenFolderDialog openFolderDialog = new();

        bool? dialog = openFolderDialog.ShowDialog();

        if (dialog != null && dialog.Value)
        {
            string outPath = openFolderDialog.FolderName;
            DateOnly dateNow = DateOnly.FromDateTime(DateTime.Now);

            string fullAdd = "Jose O. Lara<br>1234 Main St.<br>Anytown, USA 12345";

            Cartas cartas = new()
            {
                DatePrint = dateNow.ToString(),
                NPI = "1234567890",
                ICN = "ICN-1234567890",
                ServiceDate = dateNow.ToString(),
                FullAdd = fullAdd,
                IsDummy = false.ToString()
            };

            DataTable cartasToTable = _tool.ModelToTable(cartas);

            await reportService.RunWithOutTable(cartaFilePage1, cartaFilePage2, cartasToTable, outPath,
                "Demo.pdf");
        }
    }
}