using System.Data;
using System.Reflection;

namespace ReportingToPdf.tools;

public class ToolsService : IToolsService
{
    public DataTable ModelToTable<T>(T model)
    {
        DataTable dt = new();

        foreach (PropertyInfo info in typeof(T).GetProperties())
            dt.Columns.Add(new DataColumn(info.Name, info.PropertyType));

        DataRow row = dt.NewRow();

        foreach (PropertyInfo info in typeof(T).GetProperties()) row[info.Name] = info.GetValue(model, null);

        dt.Rows.Add(row);

        return dt;
    }

    public DataRow ModelToRow<T>(T model)
    {
        DataTable dt = new();

        foreach (PropertyInfo info in typeof(T).GetProperties())
            dt.Columns.Add(new DataColumn(info.Name, info.PropertyType));

        DataRow row = dt.NewRow();

        foreach (PropertyInfo info in typeof(T).GetProperties()) row[info.Name] = info.GetValue(model, null);

        //dt.Rows.Add(row);

        return row;
    }
}