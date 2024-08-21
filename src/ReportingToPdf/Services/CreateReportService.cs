using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Telerik.Reporting;
using Telerik.Reporting.Processing;
using DetailSection = Telerik.Reporting.DetailSection;
using Report = Telerik.Reporting.Report;
using SubReport = Telerik.Reporting.SubReport;
using Table = Telerik.Reporting.Table;

namespace ReportingToPdf.Services;

public class CreateReportService : ICreateReportService
{
    public async Task<bool> RunWithTable(string uriReportPath, DataTable dataTable, string fileOutPut,
        string? prefijo = null)
    {
        try
        {
            ReportProcessor reportProcessor = new();
            Hashtable deviceInfo = new();

            UriReportSource uriReportSource = new()
            {
                Uri = uriReportPath
            };

            ReportPackager reportPackager = new();
            await using FileStream sourceStream = File.OpenRead($"{uriReportSource.Uri}");
            Report reportt = reportPackager.Unpackage(sourceStream);
            DetailSection detail = (DetailSection)reportt.Items["detailSection1"];
            Table table = (Table)detail.Items["table1"];
            table.DataSource = dataTable;

            InstanceReportSource instanceReportSource = new()
            {
                ReportDocument = reportt
            };

            RenderingResult result =
                reportProcessor.RenderReport("PDF", instanceReportSource, deviceInfo);

            string fileName = $"{result.DocumentName}.{result.Extension}";

            if (!string.IsNullOrEmpty(prefijo))
                fileName = $"{prefijo}.{result.DocumentName}.{result.Extension}";

            string filePath = Path.Combine(fileOutPut, fileName);

            await using FileStream fs = new(filePath, FileMode.Create);
            //await fs.WriteAsync(result.DocumentBytes, 0, result.DocumentBytes.Length);
            await fs.WriteAsync(result.DocumentBytes.AsMemory(0, result.DocumentBytes.Length));
            await fs.FlushAsync();
            return true;
        }
        catch (Exception e)
        {
            throw new Exception(e.Message);
        }
    }

    public async Task<bool> RunWithOutTable(string uriReportPath, string uriSubReport, DataTable dataTable,
        string fileOutPut, string? fileName = null,
        string? prefijo = null)
    {
        try
        {
            ReportProcessor reportProcessor = new();
            Hashtable deviceInfo = new();

            UriReportSource uriReportSource = new()
            {
                Uri = uriReportPath
            };

            ReportPackager reportPackager = new();
            await using FileStream sourceStream = File.OpenRead($"{uriReportSource.Uri}");
            Report reportt = reportPackager.Unpackage(sourceStream);

            // Subreport
            UriReportSource subReportSource = new()
            {
                Uri = uriSubReport
            };

            // Assuming the subreport is embedded in the main report
            SubReport subReport = (SubReport)reportt.Items.Find("subReport1", true)[0];
            subReport.ReportSource = subReportSource;

            reportt.DataSource = dataTable;

            InstanceReportSource instanceReportSource = new()
            {
                ReportDocument = reportt
            };

            RenderingResult result =
                reportProcessor.RenderReport("PDF", instanceReportSource, deviceInfo);

            if (string.IsNullOrEmpty(fileName))
                fileName = $"{result.DocumentName}.{result.Extension}";

            if (!string.IsNullOrEmpty(prefijo))
                fileName = $"{prefijo}.{result.DocumentName}.{result.Extension}";

            string filePath = Path.Combine(fileOutPut, fileName);

            await using FileStream fs = new(filePath, FileMode.Create);
            //await fs.WriteAsync(result.DocumentBytes, 0, result.DocumentBytes.Length);
            await fs.WriteAsync(result.DocumentBytes.AsMemory(0, result.DocumentBytes.Length));
            await fs.FlushAsync();
            fs.Close();

            string pdfCartaPath = Path.Combine(fileOutPut, "pdfCarta\\");

            if (!Directory.Exists(pdfCartaPath))
                Directory.CreateDirectory(pdfCartaPath);

            File.Move(filePath, Path.Combine(pdfCartaPath, $"{fileName}"));

            return true;
        }
        catch (Exception e)
        {
            throw new Exception(e.Message);
        }
    }

    #region Dispose

    private IntPtr _nativeResource = Marshal.AllocHGlobal(100);

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~CreateReportService()
    {
        // Finalizer calls Dispose(false)
        Dispose(false);
    }

    private void Dispose(bool disposing)
    {
        if (disposing)
            // free managed resources AnotherResource 
            GC.Collect();

        // free native resources if there are any.
        if (_nativeResource == IntPtr.Zero) return;
        Marshal.FreeHGlobal(_nativeResource);
        _nativeResource = IntPtr.Zero;
    }

    #endregion
}