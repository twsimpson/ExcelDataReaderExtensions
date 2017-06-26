using System;
using System.Data;
using System.Linq;
using Excel;

namespace ExcelDataReaderExtensions
{

    public sealed class DataTableConfig
    {
        private static readonly string _defaultColumnNameFormat = "Column_{0}";
        private bool _useHeaderRow = true;
        private bool _detectColumnDataTypes = true;
        private string _columnNameFormat = _defaultColumnNameFormat;
        private HeaderRowSelectDelegate _headerRowSelect = _defaultHeaderRowSelect;
        private HeaderSelectDelegate _headerSelect = _defaultHeaderSelect;
        private RowSelectDelegate _rowSelect = _defaultRowSelect;
        private ColumnTypeSelectDelegate _columnTypeSelect = _defaultColumnTypeSelect;

        /// <summary>
        /// Header row is extracted if this is true (Default: true)
        /// </summary>
        public bool UseHeaderRow
        {
            get => _useHeaderRow;
            set => _useHeaderRow = value;
        }

        /// <summary>
        /// Column data is scanned to guess the data types if this is true (Default: true)
        /// </summary>
		public bool DetectColumnDataTypes
        {
            get => _detectColumnDataTypes;
            set => _detectColumnDataTypes = value;
        }

        /// <summary>
        /// Format used for column names when no header is available. (Default: "Column_{0}")
        /// Format parameter {0} is replaced with a number to ensure uniqueness.
        /// </summary>
        public string ColumnNameFomat
        {
            get => _columnNameFormat;
            set {
                if (string.IsNullOrWhiteSpace(value))
                {
                    _columnNameFormat = _defaultColumnNameFormat;
                    return;
                }
                if (!value.Contains("{0}"))
                {
                    value += "{0}";
                }
                _columnNameFormat = value;
            }
        }

        public delegate void HeaderRowSelectDelegate(ExcelOpenXmlReader reader, DataTableConfig config);

        public HeaderRowSelectDelegate HeaderRowSelect
        {
            get => _headerRowSelect;
            set => _headerRowSelect = value ?? _defaultHeaderRowSelect;
        }

        public delegate string[] HeaderSelectDelegate(ExcelOpenXmlReader reader, DataTableConfig config);

		public HeaderSelectDelegate HeaderSelect
        {
            get => _headerSelect;
            set => _headerSelect = value ?? _defaultHeaderSelect;
        }

        public delegate void RowSelectDelegate(ExcelOpenXmlReader reader, DataTableConfig config, DataRow row);

        public RowSelectDelegate RowSelect
        {
            get => _rowSelect;
            set => _rowSelect = value ?? _defaultRowSelect;
        }

        public delegate Type ColumnTypeSelectDelegate(DataColumn column, DataTableConfig config);

		public ColumnTypeSelectDelegate ColumnTypeSelect
		{
			get => _columnTypeSelect;
			set => _columnTypeSelect = value ?? _defaultColumnTypeSelect;
		}

        private static bool HasNonNullField(ExcelOpenXmlReader reader)
        {
            ExcelDataReaderExtensions.Log("Checking row for non-null fields");
            // For each field in the current row
            for (int idx = 0; idx < reader.FieldCount; idx++)
            {
                // If field is not null
                if (!reader.IsDBNull(idx))
                {
                    return true;
                }
            }
            return false;
        }
 
		private static void _defaultHeaderRowSelect(ExcelOpenXmlReader reader, DataTableConfig config)
		{
            // Select first non-null row
            // For each row, loop until found a row with a non-null value
            ExcelDataReaderExtensions.Log("Finding row with non-null fields");
            while (!HasNonNullField(reader) && reader.Read())
            {
                // Loop
            }
		}

		private static string[] _defaultHeaderSelect(ExcelOpenXmlReader reader, DataTableConfig config)
		{
            ExcelDataReaderExtensions.Log("Selecting Headers");
            var result = new string[reader.FieldCount];
			for (int i = 0; i < reader.FieldCount; i++)
			{
                var name = reader.GetString(i);
                ExcelDataReaderExtensions.Log($"Header at index {i} is {name}");
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = string.Format(config.ColumnNameFomat, i);
                }

                while (result.Contains(name))
                {
                    name = string.Format(name + $"_{i}");
                }
                ExcelDataReaderExtensions.Log($"Setting header {i} to {name}");
                result[i] = name;
			}
			return result;
		}

        private static void _defaultRowSelect(ExcelOpenXmlReader reader, DataTableConfig config, DataRow row)
        {
            // ExcelDataReaderExtensions.Log("Selecting Row"); // <- NOISY
            for (int i = 0; i < row.Table.Columns.Count; i++)
            {
                if (reader.IsDBNull(i))
                {
                    continue;
                }
                row[i] = reader.GetValue(i);
            }
        }

		private static Type _defaultColumnTypeSelect(DataColumn column, DataTableConfig config)
		{
            ExcelDataReaderExtensions.Log($"Selecting column type for column {column.ColumnName}");
            Type col_type = null;
            foreach (DataRow row in column.Table.Rows)
            {
                if (!row.IsNull(column))
                {
                    col_type = row[column].GetType();
                }
            }
			return col_type;
		}

	}
}
