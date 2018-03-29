using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _5QDataExtractor.Library.Model
{
    public class ColDTModel
    {
        // default column type (cell.Value)
        public string OldDataType { get; set; }

        // new column data type selected by the user
        public string ColumnDataType { get; set; }

        public bool IsAggregationKey { get; set; }

        // operation to perform on the column's rows
        public string RowsOperation { get; set; }

        // dictionary with mapping between distinct Excell table column cell values
        // and a user replacement value validated to convert to the new column data type
        // unique cell value: {unique cell value as str, user value}
        public Dictionary<object, List<object>> ExCellValToUserValDict { get; set; }


        public ColDTModel(string oldDT, string colDataType, bool isAggKey,
                           string rowsOperation, Dictionary<object, List<object>> ExCellValToUserValDict)
        {
            OldDataType = oldDT;

            ColumnDataType = colDataType;

            IsAggregationKey = isAggKey;

            RowsOperation = rowsOperation;

            this.ExCellValToUserValDict = ExCellValToUserValDict;
        }

        //public ColDTModel(string oldDT, string colDataType, bool isAggKey,
        //                  string rowsOperation)
        //{
        //    OldDataType = oldDT;

        //    ColumnDataType = colDataType;

        //    IsAggregationKey = isAggKey;

        //    RowsOperation = rowsOperation;

        //}
    }
}
