using System;
using Excel;

namespace ExcelDataReaderExtensions
{
    public sealed class DataSetConfig
	{

		private TableSelectDelegate _tableSelect = _defaultTableSelect;
		private DataTableConfigSelectDelegate _dataTableConfigSelect = _defaultDataTableConfigSelect;

        public delegate void TableSelectDelegate(ExcelOpenXmlReader reader, DataSetConfig config);

		public TableSelectDelegate TableSelect
		{
			get => _tableSelect;
			set => _tableSelect = value ?? _defaultTableSelect;
		}

        public delegate DataTableConfig DataTableConfigSelectDelegate(ExcelOpenXmlReader reader, DataSetConfig config);

		public DataTableConfigSelectDelegate DataTableConfigSelect
		{
			get => _dataTableConfigSelect;
			set => _dataTableConfigSelect = value ?? _defaultDataTableConfigSelect;
		}

		private static void _defaultTableSelect(ExcelOpenXmlReader reader, DataSetConfig config)
		{
            // Select first non-null table
            ExcelDataReaderExtensions.Log("Selecting non-null table");
            do
            {
                reader.NextResult();
            } while (reader.FieldCount < 1);
		}

		private static DataTableConfig _defaultDataTableConfigSelect(ExcelOpenXmlReader reader, DataSetConfig config)
		{
            ExcelDataReaderExtensions.Log("Generating default DataTableConfig");
            return new DataTableConfig();
		}

	}
}
