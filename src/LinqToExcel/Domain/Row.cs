namespace LinqToExcel
{
    using System;
    using System.Collections.Generic;
    using LinqToExcel.Extensions;
    using LinqToExcel.Query;

    public class Row : List<Cell>
    {
        IDictionary<string, int> _columnIndexMapping;
        IDictionary<string, string> ColumnNameMapping { get; } = new Dictionary<string, string>();

        public Row() :
            this(new List<Cell>(), new Dictionary<string, int>())
        { }

        /// <param name="cells">Cells contained within the row</param>
        /// <param name="columnIndexMapping">Column name to cell index mapping</param>
        public Row(IList<Cell> cells, IDictionary<string, int> columnIndexMapping)
        {
            for (int i = 0; i < cells.Count; i++)
                this.Insert(i, cells[i]);
            _columnIndexMapping = columnIndexMapping;
            foreach (string key in columnIndexMapping.Keys)
            {
                ColumnNameMapping.Add(ExcelUtilities.NormalizeColumnName(key), key);
            }
        }

        /// <param name="columnName">Column Name</param>
        public Cell this[string columnName]
        {
            get
            {
                string normalized = ExcelUtilities.NormalizeColumnName(columnName);
                if (!ColumnNameMapping.ContainsKey(normalized))
                    throw new ArgumentException(string.Format("'{0}' column name does not exist. Valid column names are '{1}'",
                        columnName, string.Join("', '", _columnIndexMapping.Keys.ToArray())));
                return base[_columnIndexMapping[ColumnNameMapping[normalized]]];
            }
        }

        /// <summary>
        /// List of column names in the row object
        /// </summary>
        public IEnumerable<string> ColumnNames
        {
            get { return _columnIndexMapping.Keys; }
        }
    }
}
