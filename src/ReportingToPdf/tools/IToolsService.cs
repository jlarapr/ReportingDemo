using System.Data;

namespace ReportingToPdf.tools;

public interface IToolsService
{
    DataRow ModelToRow<T>(T model);
    DataTable ModelToTable<T>(T model);
}