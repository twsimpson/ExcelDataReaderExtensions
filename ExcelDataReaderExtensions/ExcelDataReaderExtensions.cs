using System;
using System.Data;
using Excel;

namespace ExcelDataReaderExtensions
{
	public static class ExcelDataReaderExtensions
	{
		public static DataTable ToDataTable(this ExcelOpenXmlReader self, DataTableConfig config = null) //DONE
		{
            Log($"Extracting {self.Name} to DataTable");
            Log(@"config: {{
    UseHeaderRow = {0},
    DetectColumnTypes = {1},
    ColumnNameFormat = {2},
    HeaderRowSelect = {3},
    HeaderSelect = {4},
    RowSelect = {5},
    ColumnTypeSelect = {6}
}}",
        config.UseHeaderRow, config.DetectColumnDataTypes, config.ColumnNameFomat,
        string.Join(".", config.HeaderRowSelect.Method.DeclaringType?.Namespace, config.HeaderRowSelect.Method.DeclaringType?.Name, config.HeaderRowSelect.Method.Name).TrimStart('.'),
        string.Join(".", config.HeaderSelect.Method.DeclaringType?.Namespace, config.HeaderSelect.Method.DeclaringType?.Name, config.HeaderSelect.Method.Name).TrimStart('.'),
        string.Join(".", config.RowSelect.Method.DeclaringType?.Namespace, config.RowSelect.Method.DeclaringType?.Name, config.RowSelect.Method.Name).TrimStart('.'),
        string.Join(".", config.ColumnTypeSelect.Method.DeclaringType?.Namespace, config.ColumnTypeSelect.Method.DeclaringType?.Name, config.ColumnTypeSelect.Method.Name).TrimStart('.')
    );
            if (config == null)
			{
                Log("Using default DataTableConfig");
                config = new DataTableConfig();
			}

			var table = new DataTable { TableName = self.Name };

            self.Read(); // Position at first row
            Log($"UseHeaderRow = {config.UseHeaderRow}");
            if (config.UseHeaderRow)
			{
                Log("Extracting Headers");
                config.HeaderRowSelect(self, config);
                foreach (var header in config.HeaderSelect(self, config))
                {
                    table.Columns.Add(header);
                }
                self.Read(); // skip header row
			} else {
                Log($"Generating {self.FieldCount} header(s)");
                for (int i = 0; i < self.FieldCount; i++)
                {
                    table.Columns.Add(string.Format(config.ColumnNameFomat, i));
                }
			}

            table.BeginLoadData();
            Log("Begin reading Rows");
            do
            {
                var row = table.NewRow();
                try
                {
                    config.RowSelect(self, config, row);
                }
                catch
                {
                    break;
                }
                table.Rows.Add(row);
            } while (self.Read());
            Log($"End reading {table.Rows.Count} Row(s)");
            table.EndLoadData();

            Log($"DetectColumnTypes = {config.DetectColumnDataTypes}");
            if (config.DetectColumnDataTypes)
			{
                Log("Detecting Column data types");
                DataTable new_table = null;
                for (int colIdx = 0; colIdx < table.Columns.Count; colIdx++)
                {
                    var column = table.Columns[colIdx];
                    Log($"Checking data type for Column[{colIdx}]({column.ColumnName})");
                    var col_type = column.DataType;
                    var new_type = config.ColumnTypeSelect(column, config) ?? col_type;
                    if (col_type != new_type)
                    {
                        Log($"Found new data type for Column[{colIdx}]({column.ColumnName}) {col_type.Name} => {new_type.Name}");
                        new_table = new_table ?? table.Clone();
                        new_table.Columns[colIdx].DataType = new_type;
                    }
                }
                Log("Data types changed: {0}", new_table != null);
                if (new_table != null)
                {
                    Log($"Correcting data types for Table({table.TableName})");
                    new_table.BeginLoadData();
                    foreach (DataRow row in table.Rows)
                    {
                        new_table.ImportRow(row);
                    }
                    new_table.EndLoadData();
                    table.Dispose();
                    table = new_table;
                }
			}
            Log($"Findished Extracting {self.Name} to DataTable({table.TableName})");
            return table;
		}

		public static DataSet ToDataSet(this ExcelOpenXmlReader self, DataSetConfig config) //DONE
		{
            Log("Extracting Sheets to DataSet");
            Log(@"config: {{
    TableSelect = {0},
    DataTableConfigSelect = {1}
}}",
        string.Join(".", config.TableSelect.Method.DeclaringType?.Namespace, config.TableSelect.Method.DeclaringType?.Name, config.TableSelect.Method.Name).TrimStart('.'),
        string.Join(".", config.DataTableConfigSelect.Method.DeclaringType?.Namespace, config.DataTableConfigSelect.Method.DeclaringType?.Name, config.DataTableConfigSelect.Method.Name).TrimStart('.')
    );
            var result = new DataSet();

            do
            {
                try
                {
                    config.TableSelect(self, config);
                }
                catch
                {
                    break;
                }
                result.Tables.Add(ToDataTable(self, config.DataTableConfigSelect(self, config)));

            } while (self.NextResult());
            result.AcceptChanges();
            Log($"Finished Extracting {result.Tables.Count} Sheet(s) to DataSet");
            return result;
		}

		public static DataSet ToDataSetExt(this ExcelOpenXmlReader self, DataSetConfig config = null) => ToDataSet(self, config ?? new DataSetConfig());

        internal static void Log(string msg, params object[] args)
        {
#if DEBUG
            if (args?.Length > 0)
            {
                msg = string.Format(msg, args);
            }
            Console.WriteLine(string.Format("[{0}] ", DateTime.Now) + msg);
#else
#endif
        }

    }
}

