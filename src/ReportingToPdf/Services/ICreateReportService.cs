using System.Data;

namespace ReportingToPdf.Services;

public interface ICreateReportService : IDisposable
{
    Task<bool> RunWithOutTable(string uriReportPath, string uriSubReport, DataTable dataTable, string fileOutPut,
        string? prefijo = null, string? fileName = null);

    Task<bool> RunWithTable(string uriReportPath, DataTable dataTable, string fileOutPut, string? prefijo = null);
}